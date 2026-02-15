on run {outputPath, pptPath}
    tell application "Microsoft PowerPoint"
        -- activate removed
        
        -- CHECK IF ALREADY OPEN
        set pres to missing value
        try
            repeat with p in presentations
                if full name of p contains pptPath then
                    set pres to p
                    exit repeat
                end if
            end repeat
        end try
        
        if pres is missing value then
             if pptPath is not "" then
                open (POSIX file pptPath)
                set pres to active presentation
             else
                -- If no path provided, assume active active
                if exists active presentation then
                    set pres to active presentation
                end if
             end if
        end if
        
        if pres is missing value then
            return "Error: No active presentation to export."
        end if
        
        try
            -- Export using the specific presentation object
            save pres in (POSIX file outputPath) as save as movie
            
            -- 3. WAIT FOR COMPLETION happens below
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
    end repeat
    
    if isBusy then
         return "Error: Timeout. PowerPoint is still holding the file open."
    end if
    
    -- 4. DO NOT CLOSE
    -- The user wants the presentation to stay open.
    
    return "Success"
end run
