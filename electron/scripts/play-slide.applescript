on run {slideIndex}
	set slideIndex to slideIndex as string
	
	-- 1. Write Slide Index to File
	-- This file is read by the VBA macro 'PlaySlide'
	set paramPath to (path to library folder from user domain as string) & "Group Containers:UBF8T346G9.Office:play_slide.txt"
	set posixParamPath to POSIX path of paramPath
	
	try
		do shell script "echo " & slideIndex & " > " & quoted form of posixParamPath
	on error errMsg
		return "Error writing play_slide.txt: " & errMsg
	end try
	
	-- 2. Trigger the Macro
	tell application "Microsoft PowerPoint"
		activate
		try
			-- Try running with the specific add-in syntax
			run VB macro macro name "AudioTools.ppam!PlaySlide"
		on error errMsg1
			try
				-- Fallback to simple name (if imported as module)
				run VB macro macro name "PlaySlide"
			on error errMsg2
				return "Error running macro: " & errMsg1 & " // " & errMsg2
			end try
		end try
	end tell
	
end run
