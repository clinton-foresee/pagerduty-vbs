# Overview
This VBScript triggers and resolves incidents in PagerDuty when an event is added to a queue directory (`C:\PagerDuty\Queue`).

## Example Usage
SolarWinds alerts are logged to a text file (which contain data in PagerDuty's [Events API JSON format](https://v2.developer.pagerduty.com/docs/events-api)) in the queue directory, and the VBScript sends the alert files in the queue directory to PagerDuty via a follow-up action from SolarWinds or Windows scheduled task that executes the VBScript with [CScript.exe](https://technet.microsoft.com/en-us/library/bb490887.aspx).

There are example alert triggers and resets for SolarWinds NPM and SAM in the **SolarWinds Sample Alerts** directory. Importing these will start logging alerts to `C:\PagerDuty\Queue`.

For a full walkthrough on setting up the integration, please see the [SolarWinds Integration Guide](https://www.pagerduty.com/docs/guides/solarwinds-integration-guide/) on pagerduty.com.