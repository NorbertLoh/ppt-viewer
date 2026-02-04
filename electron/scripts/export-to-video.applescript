on run {outputPath, pptPath}
    tell application "Microsoft PowerPoint"
        -- activate removed
        
        -- Ensure the correct presentation is active
        if pptPath is not "" then
            open (POSIX file pptPath)
        end if
        
        if not (exists active presentation) then
            return "Error: No active presentation to export."
        end if
        
        try
            -- Converting POSIX path to Mac path usually required for 'save as'
            -- But "save active presentation in (POSIX file ...)" works in newer versions.
            -- Using "save as movie" creates an MP4 or MOV depending on version/settings.
            
            save active presentation in (POSIX file outputPath) as save as movie
            
            -- 3. WAIT FOR COMPLETION
            -- Loop checking file size until it stops changing
        on error errMsg
            return "Error initiating export: " & errMsg
        end try
        
    end tell
    
    -- OUTSIDE PPT BLOCK: Check file system using Shell (reliable)
    
    set stableCount to 0
    set lastSize to -1
    set waitCycles to 0
    set maxInitialWait to 30 -- Wait up to 30s for file to APPEAR
    set maxStableWait to 600 -- Wait up to 10 mins for export
    
    -- 1. Wait for file to be created
    repeat while waitCycles < maxInitialWait
        try
            do shell script "ls " & quoted form of outputPath
            exit repeat -- File exists
        on error
            -- File not found yet
        end try
        delay 1
        set waitCycles to waitCycles + 1
    end repeat
    
    if waitCycles >= maxInitialWait then
        return "Error: Timeout waiting for video file to be created."
    end if
    
    -- 2. Wait for file to be released (lsof check)
    -- If lsof finds the file open, it returns 0 (success). If not open, it returns 1 (error).
    set isBusy to true
    set waitCycles to 0
    
    repeat while isBusy and waitCycles < maxStableWait
        delay 2
        set waitCycles to waitCycles + 2
        
        try
            -- Check if anyone (specifically PPT) has the file open
            do shell script "lsof " & quoted form of outputPath
            -- If we are here, grep found it (exit code 0), so it IS busy
            set isBusy to true
        on error
            -- lsof failed (exit code 1), meaning file is NOT open by anyone
            set isBusy to false
        end try
        
        -- Logging for debug (optional, but good if we could see it)
        -- do shell script "echo 'Busy: " & isBusy & "' >&2"
    end repeat
    
    if isBusy then
         return "Error: Timeout. PowerPoint is still holding the file open."
    end if
    
    -- 4. CLOSE
    tell application "Microsoft PowerPoint"
        try
             close active presentation saving no
            
            return "Success"
        on error errMsg
            return "Error exporting video: " & errMsg
        end try
    end tell
end run
