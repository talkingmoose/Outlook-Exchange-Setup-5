#!/bin/sh

# Check for the existence of the UBF8T346G9.Office folder.
# If it doesn't exist then no Office 2016 for Mac application has run.
# Create the folder create the OutlookProfile.plist file.
# Also create a user LaunchAgents folder and an Outlook setup launchd agent.

if [[ ! -d "$HOME/Library/Group Containers/UBF8T346G9.Office" ]] ; then
	mkdir -p "$HOME/Library/Group Containers/UBF8T346G9.Office"
	touch "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist"
	mkdir -p "$HOME/Library/LaunchAgents"

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
	<string>net.talkingmoose.OutlookExchangeSetup5.0</string>
	<key>ProgramArguments</key>
	<array>
		<string>/usr/bin/osascript</string>
		<string>/Library/Talking Moose Industries/Scripts/Outlook Exchange Setup 5.0.scpt</string>
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

	echo "$launchagent" >> "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.0.plist"
	chmod 644 "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.0.plist"
	launchctl load "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.0.plist"
	
fi

# Check for the existence of anything in an Outlook profile's Exchange Accounts folder.
# If anything exists, an Exchange account exists. Nothing more to do.
# Unload and remove the launch agent if it stil exists.

profile=$( defaults read "$HOME/Library/Group Containers/UBF8T346G9.Office/OutlookProfile.plist" Default_Profile_Name )

if [[ -f "$HOME/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/$profile/Data/Exchange Accounts/"* ]] ; then
	launchctl unload "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.0.plist"
	rm "$HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.0.plist"
fi

exit 0