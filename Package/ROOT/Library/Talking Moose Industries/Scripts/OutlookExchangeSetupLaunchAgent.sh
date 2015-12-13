#!/bin/sh

# --------------------------------------------
# Outlook Exchange Setup 5.0
# © Copyright 2008-2015 William Smith
# bill@officeformachelp.com
# 
# Except where otherwise noted, this work is licensed under
# http://creativecommons.org/licenses/by/4.0/
#
# This file is one of four files for assisting a user with configuring
# an Exchange account in Microsoft Outlook 2016 for Mac:
# 
# 1. Outlook Exchange Setup 5.1.0.scpt
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

# Check for the existence of the UBF8T346G9.Office folder.
# If it doesn't exist then no Office 2016 for Mac application has run.
# Create the folder create the OutlookProfile.plist file.
# Also create a user LaunchAgents folder and an Outlook setup launchd agent.

if [[ ! -d "$HOME/Library/Group Containers/UBF8T346G9.Office" ]] ; then
	/bin/mkdir -p "$HOME/Library/Group Containers/UBF8T346G9.Office"
	/usr/bin/touch "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist"
	/bin/mkdir -p "$HOME/Library/LaunchAgents"

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
	<string>net.talkingmoose.OutlookExchangeSetup5.1.0</string>
	<key>ProgramArguments</key>
	<array>
		<string>/usr/bin/osascript</string>
		<string>/Library/Talking Moose Industries/Scripts/Outlook Exchange Setup 5.1.0.scpt</string>
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

	/bin/echo "$launchagent" >> "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.1.0.plist"
	/bin/chmod 644 "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.1.0.plist"
	/bin/launchctl load "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.1.0.plist"
	
fi

# Check for the existence of anything in an Outlook profile's Exchange Accounts folder.
# If anything exists, an Exchange account exists. Nothing more to do.
# Unload and remove the launch agent if it stil exists.

profile=$( /usr/bin/defaults read "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist" Default_Profile_Name )

if [[ -f "$HOME/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/$profile/Data/Exchange Accounts/"* ]] ; then
	/bin/launchctl unload "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.1.0.plist"
	/bin/rm "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.1.0.plist"
fi

exit 0