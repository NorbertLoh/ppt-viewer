on run {audioPath, slideIndex, pptPath}
    -- Path to the parameters file in the Office Group Container
    -- This path is shared between Electron (outside sandbox) and PowerPoint (sandboxed)
    set homePath to POSIX path of (path to home folder)
    set paramsInfoPath to homePath & "Library/Group Containers/UBF8T346G9.Office/audio_params.txt"
    
    -- Content to write: "SlideIndex|AudioPath|PresentationPath"
    set fileContent to (slideIndex as text) & "|" & audioPath & "|" & pptPath
    
    -- Write to file using shell command (reliable)
    do shell script "echo " & quoted form of fileContent & " > " & quoted form of paramsInfoPath
    
    tell application "Microsoft PowerPoint"
        -- activate removed
        open (POSIX file pptPath)
        
        try
            -- Call the macro without arguments
            run VB macro macro name "InsertAudio"
        on error errMsg
            return "Error calling macro: " & errMsg
        end try
        
    end tell
end run
