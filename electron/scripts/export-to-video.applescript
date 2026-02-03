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
    
    -- 2. Wait for stability
    set waitCycles to 0
    repeat while stableCount < 5 and waitCycles < maxStableWait
        delay 1
        set waitCycles to waitCycles + 1
        
        try
            -- Get file size in bytes using stat (MacOS bsd stat)
            set curSize to (do shell script "stat -f%z " & quoted form of outputPath) as integer
        on error
            set curSize to -1
        end try
        
        if curSize > 0 then
            if curSize = lastSize then
                set stableCount to stableCount + 1
            else
                set stableCount to 0
                set lastSize to curSize
            end if
        else
            set stableCount to 0
        end if
    end repeat
    
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
