## What is it?
This is a powershell script to parse the iCal feed from Who's Off (www.whosoff.com), and publish it to a Slack channel or user.

## Getting started
- Create a feed in Who's Off, and copy the feed URL (i.e. `http://feeds.staff.whosoff.com/?u=0000000000`)
- Create an incoming webhook in Slack. Copy the URL (i.e. `https://hooks.slack.com/services/A000000C/B000000U/SAKLJDKSALJDKLSAJD`)
- Amend the command line example below, to use your URLs, and replace the -slackchannel parameter with your channel or username.

## Example usage at the Poweshell command line
`.\whosoff.ps1 -whosoffurl http://feeds.staff.whosoff.com/?u=0000000000 -slackchannel #yourchannel -slackapihookurl https://hooks.slack.com/services/A000000C/B000000U/SAKLJDKSALJDKLSAJD`

## Example usage using Task Scheduler
`C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File "c:\PowerShell scripts\whosoff.ps1" -whosoffurl "http://feeds.staff.whosoff.com/?u=0000000000" -slackchannel "#yourchannel" -slackapihookurl "https://hooks.slack.com/services/A000000C/B000000U/SAKLJDKSALJDKLSAJD"`
