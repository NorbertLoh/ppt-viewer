on run {macroName, pptPath}
    tell application "Microsoft PowerPoint"
        -- activate removed
        
        -- open command usually activates, so we might want to check if open first?
        -- But for now standard open is fine as long as we don't force activate.
        open (POSIX file pptPath)
        
        try
            -- Run the specified macro
            run VB macro macro name macroName
        on error errMsg
            return "Error calling macro '" & macroName & "': " & errMsg
        end try
        
    end tell
end run
