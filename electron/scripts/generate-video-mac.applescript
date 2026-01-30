on run argv
    -- Argument Handling
    if count of argv is less than 3 then
        error "Usage: generate-video-mac.applescript <inputPath> <audioDir> <outputPath>"
    end if
    
    set inputPath to item 1 of argv
    set audioDir to item 2 of argv
    set outputPath to item 3 of argv
    
    -- Setup Sandbox Temp Directory
    -- PowerPoint is sandboxed and can only read files freely from its own container.
    -- We copy audio files to its tmp directory to avoid permission prompts and errors.
    set userHome to (path to home folder as text)
    set runID to (do shell script "date +%s")
    -- HFS Path to Sandbox Tmp: Macintosh HD:Users:name:Library:Containers:com.microsoft.Powerpoint:Data:tmp:ppt_gen_ID
    set sandboxRelativePath to "Library:Containers:com.microsoft.Powerpoint:Data:tmp:ppt_gen_" & runID
    set sandboxTmpHFS to userHome & sandboxRelativePath
    set sandboxTmpPOSIX to POSIX path of sandboxTmpHFS
    
    do shell script "mkdir -p " & quoted form of sandboxTmpPOSIX
    
    tell application "Microsoft PowerPoint"
        activate
        
        -- Open the presentation
        try
            open (POSIX file inputPath)
            set pptPres to active presentation
        on error errMsg
            -- Cleanup on failure
            try
                do shell script "rm -rf " & quoted form of sandboxTmpPOSIX
            end try
            error "Failed to open presentation: " & errMsg
        end try
        
        -- Iterate through slides
        set slideCount to count of slides of pptPres
        
        repeat with i from 1 to slideCount
            set currentSlide to slide i of pptPres
            set sourceAudio to audioDir & "/slide_" & i & ".wav"
            
            -- Check if source audio file exists
            set audioExists to false
            try
                do shell script "test -f " & quoted form of sourceAudio
                set audioExists to true
            on error
                set audioExists to false
            end try
            
            if audioExists then
                -- Copy to Sandbox
                set destAudioPOSIX to sandboxTmpPOSIX & "/slide_" & i & ".wav"
                do shell script "cp " & quoted form of sourceAudio & " " & quoted form of destAudioPOSIX
                
                -- Construct HFS path manually to avoid coercion issues
                -- Replace slash with colon and prepend drive name (assuming default startup disk)
                -- Getting startup disk name
                tell application "System Events" to set startupDisk to name of startup disk
                set destAudioHFS to startupDisk & (do shell script "echo " & quoted form of destAudioPOSIX & " | sed 's/\\//:/g'")
                
                -- Add Audio Media Object using validated HFS path
                set mediaShape to make new media object at currentSlide with properties {file name:destAudioHFS, top:10, |left|:10, width:50, height:50, link to file:false, save with document:true}
                
                -- Configure Animation Settings
                tell animation settings of mediaShape
                     set animate to true
                     tell play settings
                          set |play on entry| to true
                          set |hide while not playing| to true
                     end tell
                end tell
                
                -- Get Audio Duration
                set audioLenSec to 3.0
                try
                    set durationStr to (do shell script "mdls -raw -name kMDItemDurationSeconds " & quoted form of sourceAudio)
                    if durationStr is not "(null)" then
                        set audioLenSec to durationStr as real
                    end if
                on error
                    set audioLenSec to 3.0
                end try
                
                -- Set Slide Transition
                tell slide show transition of currentSlide
                     set advance on time to true
                     set advance time to audioLenSec
                end tell
                
            else
                -- No Audio: Default Transition
                tell slide show transition of currentSlide
                     if advance on time is false then
                         set advance on time to true
                         set advance time to 3.0
                     end if
                end tell
            end if
        end repeat
        
        -- Export Video to Sandbox First
        -- Writing directly to external folders often fails silently due to sandbox restrictions.
        -- We export to the temp folder, then move it to the destination.
        
        -- Use .mov extension for temp file as we are using 'save as movie'
        set sandboxOutputPOSIX to sandboxTmpPOSIX & "/temp_video.mov"
        -- Construct HFS path for sandbox output
        tell application "System Events" to set startupDisk to name of startup disk
        set sandboxOutputHFS to startupDisk & (do shell script "echo " & quoted form of sandboxOutputPOSIX & " | sed 's/\\//:/g'")
        
        try
            -- Try MOV legacy format (often 'save as movie' works where MP4 fails)
             run script "tell application \"Microsoft PowerPoint\" to save active presentation in \"" & sandboxOutputHFS & "\" as save as movie"
        on error errMsg
            -- Cleanup on failure
            try
                do shell script "rm -rf " & quoted form of sandboxTmpPOSIX
            end try
            error "Failed to export video: " & errMsg
        end try
        
        -- Wait for Export Completion in Sandbox
        delay 2
        set exportComplete to false
        set previousSize to -1
        set stableCount to 0
        
        repeat 300 times
            try
                 do shell script "test -f " & quoted form of sandboxOutputPOSIX
                 set fileExists to true
            on error
                 set fileExists to false
            end try
            
            if fileExists then
                set currentSize to (do shell script "stat -f%z " & quoted form of sandboxOutputPOSIX) as integer
            else
                set currentSize to -1
            end if
            
            if currentSize > 0 then
                if currentSize is equal to previousSize then
                    set stableCount to stableCount + 1
                else
                    set stableCount to 0
                end if
                
                set previousSize to currentSize
            end if
            
            -- If size is stable for 5 seconds, assume done
            if stableCount >= 5 then
                set exportComplete to true
                exit repeat
            end if
            
            delay 1
        end repeat
        
        if not exportComplete then
             error "Video export timed out in sandbox."
        end if
        
        -- Move file to final destination
        try
            do shell script "mv " & quoted form of sandboxOutputPOSIX & " " & quoted form of outputPath
        on error moveErr
             error "Failed to move video to final destination: " & moveErr
        end try
        
        -- Close without saving changes to the PPT text
        close pptPres saving no
        
        -- Quit PowerPoint (matches the behavior of the win script)
        quit
        
    end tell
    
    -- Final Cleanup
    try
        do shell script "rm -rf " & quoted form of sandboxTmpPOSIX
    end try
end run
