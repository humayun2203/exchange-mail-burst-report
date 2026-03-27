# exchange-mail-burst-report
PowerShell script to detect mail bursts in Exchange message tracking logs
# Exchange Mail Burst Report (PowerShell)

A PowerShell script for Microsoft Exchange environments that detects possible mail bursts by analyzing message tracking logs.

## What it does

This script:

- scans Exchange message tracking logs for a configurable time window
- groups messages by **Sender + Subject**
- normalizes common subject prefixes such as `RE:`, `FW:`, and `FWD:`
- filters out configured internal sender domains
- reports only entries that exceed a defined threshold
- exports results to a CSV file for further review

## Use case

This is useful for identifying patterns such as:

- repeated inbound messages from the same sender
- suspicious message bursts
- operational spikes tied to a specific subject line
- high-volume log patterns that may deserve review

## Parameters

| Parameter | Description |
|---|---|
| `WindowMinutes` | How far back to search in the tracking logs |
| `Threshold` | Minimum number of matching messages required for reporting |
| `OutDir` | Output folder for the CSV file |
| `BaseName` | Base file name used for the export |
| `EventId` | Tracking log event type (`RECEIVE`, `SEND`, or `DELIVER`) |
| `ExcludeSenderDomains` | Sender domains to exclude from the report |

## Example

```powershell
.\MailBurstReport.ps1 `
  -WindowMinutes 1440 `
  -Threshold 20 `
  -OutDir "C:\Reports" `
  -EventId RECEIVE `
  -ExcludeSenderDomains @("internal.example")
