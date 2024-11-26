#!/bin/zsh

echo "Office-Reset: Starting postinstall for Remove_ZoomPlugin"
autoload is-at-least
APP_NAME="Zoom Outlook Plugin"

GetLoggedInUser() {
	LOGGEDIN=$(/bin/echo "show State:/Users/ConsoleUser" | /usr/sbin/scutil | /usr/bin/awk '/Name :/&&!/loginwindow/{print $3}')
	if [ "$LOGGEDIN" = "" ]; then
		echo "$USER"
	else
		echo "$LOGGEDIN"
	fi
}

SetHomeFolder() {
	HOME=$(dscl . read /Users/"$1" NFSHomeDirectory | cut -d ':' -f2 | cut -d ' ' -f2)
	if [ "$HOME" = "" ]; then
		if [ -d "/Users/$1" ]; then
			HOME="/Users/$1"
		else
			HOME=$(eval echo "~$1")
		fi
	fi
}

## Main
LoggedInUser=$(GetLoggedInUser)
SetHomeFolder "$LoggedInUser"
echo "Office-Reset: Running as: $LoggedInUser; Home Folder: $HOME"

echo "Office-Reset: Running native uninstall routine for ${APP_NAME}"
/Applications/ZoomOutlookPlugin/Uninstall/Contents/MacOS/Uninstall

echo "Office-Reset: Stopping Zoom agent"
/bin/launchctl stop /Library/LaunchAgents/us.zoom.pluginagent.plist
/bin/launchctl unload /Library/LaunchAgents/us.zoom.pluginagent.plist
/bin/launchctl stop "$HOME/Library/LaunchAgents/us.zoom.pluginagent.plist"
/bin/launchctl unload "$HOME/Library/LaunchAgents/us.zoom.pluginagent.plist"

echo "Office-Reset: Removing agent configuration for ${APP_NAME}"
/bin/rm -f "/Library/LaunchAgents/us.zoom.pluginagent.plist"
/bin/rm -f "$HOME/Library/LaunchAgents/us.zoom.pluginagent.plist"
/bin/rm -f "$HOME/Library/Logs/zoomoutlookplugin.log"
/bin/rm -f "$HOME/Library/Preferences/ZoomChat.plist"

echo "Office-Reset: Removing binaries for ${APP_NAME}"
/bin/rm -rf "/Library/Application Support/ZoomOutlookPlugin"
/bin/rm -rf "/Library/Application Support/Microsoft/ZoomOutlookPlugin"
/bin/rm -rf "/Library/ScriptingAdditions/zOLPluginInjection.osax"
/bin/rm -rf "/Users/Shared/ZoomOutlookPlugin"

/bin/rm -rf "/Applications/ZoomOutlookPlugin"

/usr/sbin/pkgutil --forget ZoomMacOutlookPlugin.pkg

exit 0