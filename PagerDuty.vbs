' Copyright (c) 2016, PagerDuty, Inc. <info@pagerduty.com>
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of PagerDuty Inc nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL PAGERDUTY INC BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


' This script sends the content of text files in the QueuePath to the URL.
'
' Text files should contain all of the JSON neccessary to trigger, acknowledge, or
' resolve an incident in PagerDuty as defined in the PagerDuty Developer documentation:
' https://v2.developer.pagerduty.com/docs/events-api

On Error Resume Next

' Constants and shell object used for logging to the Windows Application Event Log

Const EVENT_SUCCESS	= 0
Const EVENT_ERROR 	= 1
Const EVENT_WARNING = 2
Const EVENT_INFO 	= 4

Set objShell = WScript.CreateObject("WScript.Shell")

' Set the API endpoint and alert file directory we're working with, and whether
' you want successful calls to be logged in the Windows Application Event Log.
' NOTE: Trailing backslash is required for QueuePath!

Dim URL, QueuePath, LogSuccess

URL = "https://events.pagerduty.com/generic/2010-04-15/create_event.json"
QueuePath = "C:\PagerDuty\Queue\"
LogSuccess = True

' Get list of filenames in the queue directory

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(QueuePath)
Set colFiles = objFolder.Files

For Each objFile in colFiles

	' Set filename variables and check for the extension ".txt" so we don't mess
	' with lock files or anything else in the queue that isn't expected to be there.

	Dim AlertFileName, AlertFileExtension, AlertFile

	AlertFileName = objFile.Name
	AlertFileExtension = objFSO.GetExtensionName(AlertFileName)
	AlertFile = QueuePath & AlertFileName

	If AlertFileExtension = "txt" Then

		' Check for a lock and quit if we're already working on this alert in another
		' process, or create a lock if we're going to start working on this alert.

		Dim PostBody, AlertLockFile, AlertFileContent, Status, Response

		AlertLockFile = AlertFile & ".lock"

		If objFSO.FileExists(AlertLockFile) Then
			WScript.Echo "Lock file already exists: " & AlertLockFile & ". Moving on to next alert file."
		Else

			WScript.Echo "Creating lock file: " & AlertLockFile
			objFSO.CreateTextFile(AlertLockFile)

			' Open and get the alert file content, escaping backslashes with an additional backslash

			Err.Clear
			Set AlertFileContent = CreateObject("ADODB.Stream")
			AlertFileContent.CharSet = "utf-8"
			AlertFileContent.Open
			AlertFileContent.LoadFromFile(AlertFile)
			PostBody = AlertFileContent.ReadText()
			PostBody = Replace(PostBody, "\", "\\")
			PostBody = Replace(PostBody,vbCr,"")
			PostBody = Replace(PostBody,vbLf,"")
			AlertFileContent.Close

			If Err.Number <> 0 Then
				WScript.Echo "ERROR: Couldn't read alert file: " & AlertFile & vbNewLine &_
					"Check Windows Application Event Log for details."

				objShell.LogEvent EVENT_ERROR, "Couldn't read alert file." & vbNewLine & vbNewLine &_
					"File name: " & AlertFile & vbNewLine &_
					"Error Number: " & Err.Number & vbNewLine &_
					"Source: " & Err.Source & vbNewLine &_
					"Description: " & Err.Description
			Else
				' Send the alert file content to PagerDuty and check response

				Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")

				Err.Clear
				WScript.Echo "Body being sent to PagerDuty is: " & vbNewLine & PostBody
				objHTTP.Open "POST", URL, False
				objHTTP.setRequestHeader "Content-Type", "application/json"
				objHTTP.Send PostBody

				' Log connection failures in case the system isn't able to resolve
				' the domain or server isn't responding at all.

				If Err.Number <> 0 Then
					WScript.Echo "ERROR: Couldn't connect or send data to PagerDuty. Check Windows Application Event Log for details."

					objShell.LogEvent EVENT_ERROR, "Couldn't connect or send data to PagerDuty." & vbNewLine & vbNewLine &_
						"Error Number: " & Err.Number & vbNewLine &_
						"Source: " & Err.Source & vbNewLine &_
						"Description: " & Err.Description
				Else
					Status = objHTTP.Status
					Response = objHTTP.responseText
					WScript.Echo "Response from PagerDuty is: [" & Status & "] " & Response

					' Remove the alert from the queue once it has been accepted by PagerDuty,
					' or log why the event wasn't accepted by PagerDuty.

					If Status = 200 Then
						WScript.Echo "Deleting alert file: " & AlertFile

						objFSO.DeleteFile(AlertFile)

						If LogSuccess = True Then
							objShell.LogEvent EVENT_SUCCESS, "PagerDuty accepted event with data:" & vbNewLine & vbNewLine &_
								PostBody & vbNewLine & vbNewLine &_
								"Response was:" & vbNewLine & vbNewLine &_
								"[" & Status & "] " & Response
						End If
					Else
						WScript.Echo "Non-200 response received. Keeping alert file in queue: " & AlertFile

						objShell.LogEvent EVENT_ERROR, "PagerDuty did not accept event with data:" & vbNewLine & vbNewLine &_
							PostBody & vbNewLine & vbNewLine &_
							"Response was:" & vbNewLine & vbNewLine &_
							"[" & Status & "] " & Response
					End If
				End If
			End If

			' Remove the lock file. This will let us try again in case PagerDuty couldn't be reached.

			WScript.Echo "Deleting lock file: " & AlertLockFile
			objFSO.DeleteFile(AlertLockFile)
		End If
	End If
Next
