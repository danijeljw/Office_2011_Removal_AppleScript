tell application "Microsoft Database Daemon" to quit
tell application "Finder"
	delete file "Microsoft Communicator.app" of (path to applications folder)
	delete file "Microsoft Messenger.app" of (path to applications folder)
	delete folder "Microsoft Office 2011" of (path to applications folder)
	delete file "Remote Desktop Connection.app" of (path to applications folder)
	delete (every item of folder "Automator" of (path to library folder) ¬
		whose name contains "Excel")
	delete (every item of folder "Automator" of (path to library folder) ¬
		whose name contains "Office")
	delete (every item of folder "Automator" of (path to library folder) ¬
		whose name contains "Outlook")
	delete (every item of folder "Automator" of (path to library folder) ¬
		whose name contains "PowerPoint")
	delete (every item of folder "Automator" of (path to library folder) ¬
		whose name contains "Word")
	delete file "Add New Sheet to Workbooks.action" of folder ¬
		"Automator" of (path to library folder)
	delete file "Create List from Data in Workbook.action" of folder ¬
		"Automator" of (path to library folder)
	delete file "Create Table from Data in Workbook.action" of folder ¬
		"Automator" of (path to library folder)
	delete file "Get Parent Presentations of Slides.action" of folder ¬
		"Automator" of (path to library folder)
	delete file "Get Parent Workbooks.action" of folder ¬
		"Automator" of (path to library folder)
	delete file "Set Document Settings.action" of folder ¬
		"Automator" of (path to library folder)
	delete folder "Microsoft" of (path to fonts folder from local domain)
	delete (every item of folder "Internet Plug-Ins" of (path to library folder) ¬
		whose name contains "SharePoint")
	delete (every item of folder "LaunchDaemons" of (path to library folder) ¬
		whose name contains "microsoft")
	delete (every item of folder "Preferences" of (path to library folder) ¬
		whose name contains "microsoft")
	delete (every item of folder "PrivilegedHelperTools" of (path to library folder) ¬
		whose name contains "microsoft")
	set receiptsList to do shell script "pkgutil --pkgs=com.microsoft.office*"
	repeat with receiptNumber from 1 to count of paragraphs in receiptsList
		set forget to "pkgutil --forget " & paragraph receiptNumber of receiptsList as string
		do shell script forget
	end repeat
end tell
