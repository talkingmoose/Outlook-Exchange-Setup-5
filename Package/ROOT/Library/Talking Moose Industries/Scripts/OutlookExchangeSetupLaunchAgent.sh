5.5.1#!/bin/sh

# --------------------------------------------
# Outlook Exchange Setup 5
# © Copyright 2008-2017 William Smith
# bill@officeformachelp.com
# 
# Except where otherwise noted, this work is licensed under
# http://creativecommons.org/licenses/by/4.0/
#
# This file is one of four files for assisting a user with configuring
# an Exchange account in Microsoft Outlook 2016 for Mac:
# 
# 1. Outlook Exchange Setup 5.5.1.scpt
# 2. OutlookExchangeSetupLaunchAgent.sh
# 3. net.talkingmoose.OutlookExchangeSetupLaunchAgent.plist
# 4. com.microsoft.Outlook.plist for creating a configuraiton profile
# 
# These scripts and files may be freely modified for personal or commercial
# purposes but may not be republished for profit without prior consent.
# 
# If you find these resources useful or have ideas for improving them,
# please let me know. It is only compatible with Outlook 2016 for Mac.
# --------------------------------------------

##### Definitions

logfile="$HOME/Library/Logs/OutlookExchangeSetup5.log"

###### Functions

function logresult()	{
	if [ $? = 0 ] ; then
	  /bin/date "+%-m/%-d/%y %-H:%M:%S %p	$1" >> "$logfile"
	else
	  /bin/date "+%-m/%-d/%y %-H:%M:%S %p	$2" >> "$logfile"
	fi
}

# Check for the existence of the UBF8T346G9.Office folder.
# If it doesn't exist then no Office 2016 for Mac application has run.
# Create the folder and create the OutlookProfile.plist file.
# Also create a user LaunchAgents folder and an Outlook setup launchd agent.

if [[ ! -d "$HOME/Library/Group Containers/UBF8T346G9.Office" ]] ; then
	logresult "Folder \"$HOME/Library/Group Containers/UBF8T346G9.Office\" does not exist."
	
	/bin/mkdir -p "$HOME/Library/Group Containers/UBF8T346G9.Office"
	logresult "Create folder \"$HOME/Library/Group Containters/UBF8T346G9.Office\": Successful." "Create folder \"$HOME/Library/Group Containters/UBF8T346G9.Office\": Failed."
	
	/usr/bin/touch "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist"
	
	logresult "Create empty file \"$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist\": Successful." "Create empty file \"$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist\": Failed."
	
	/bin/mkdir -p "$HOME/Library/LaunchAgents"
	
	logresult "Create folder \"$HOME/Library/LaunchAgents\": Successful." "Create folder \"$HOME/Library/LaunchAgents\": Failed."

	launchagent='<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>Disabled</key>
	<false/>
	<key>EnvironmentVariables</key>
	<dict>
		<key>PATH</key>
		<string>/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin:/Applications/Server.app/Contents/ServerRoot/usr/bin:/Applications/Server.app/Contents/ServerRoot/usr/sbin:/usr/local/sbin</string>
	</dict>
	<key>Label</key>
	<string>net.talkingmoose.OutlookExchangeSetup5</string>
	<key>ProgramArguments</key>
	<array>
		<string>/usr/bin/osascript</string>
		<string>/Library/Talking Moose Industries/Scripts/Outlook Exchange Setup 5.5.1.scpt</string>
	</array>
	<key>RunAtLoad</key>
	<false/>
	<key>WatchPaths</key>
	<array>
		<string>~/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist</string>
	</array>
</dict>
</plist>
'

	/bin/echo "$launchagent" >> "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist"
	
	logresult "Create launch agent \"$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist\": Successful." "Create folder \"$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist\": Failed."

	/bin/chmod 644 "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist"

	logresult "Set launch agent permissions to 644 (-rw-r--r--): Successful." "Set launch agent permissions to 644 (-rw-r--r--): Failed."

	/bin/launchctl load "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist"

	logresult "Load launch agent: Successful." "Load launch agent: Failed."
	
else
	if [[ -d "$HOME/Library/Group Containers/UBF8T346G9.Office" ]] ; then
		logresult "$HOME/Library/Group Containers/UBF8T346G9.Office folder already exists. Doing nothing." "$HOME/Library/Group Containers/UBF8T346G9.Office folder does not exist but it should exist already. Something may be wrong."
	fi
	
	if [[ -f "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist" ]] ; then
		logresult "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist already exists. Doing nothing." "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist does not exist but it should exist already. Something may be wrong."
	fi
	
	if [[ -d "$HOME/Library/LaunchAgents" ]] ; then
		logresult "$HOME/Library/LaunchAgents already exists. Doing nothing." "$HOME/Library/LaunchAgents does not exist but it should exist already. Something may be wrong."
	fi
	
	if [[ ! -f "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist" ]] ; then
		logresult "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist does not exist. Doing nothing." "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist exists. Something may be wrong."
	fi
fi

exit 0