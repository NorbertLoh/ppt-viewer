on run argv
    -- Argument Handling
    if count of argv is less than 3 then
        error "Usage: generate-video-mac.applescript <inputPath> <audioDir> <outputPath>"
    end if
    
    set inputPath to item 1 of argv
    set audioDir to item 2 of argv
    set outputPath to item 3 of argv
    
    -- Get Script Directory
    set scriptPath to POSIX path of (path to me)
    set scriptDir to do shell script "dirname " & quoted form of scriptPath
    
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
                -- Place off-screen (-100, -100) to hide the icon in the video
                set mediaShape to make new media object at currentSlide with properties {file name:destAudioHFS, top:-100, |left|:-100, width:50, height:50, link to file:false, save with document:true}
                
                -- Unique name for scripting
                set mediaName to "Audio_" & (do shell script "date +%s")
                set name of mediaShape to mediaName
                
                -- Get Slide Index for the run script context
                set sIndex to slide index of currentSlide
                
                -- STRATEGY V14: STANDARD ADD + XML BATCH
                -- We add audio with default settings (On Click).
                -- The external XML patcher will handle "Start With Previous" and "First in Timeline".
                
                -- Audio added. 
                -- We will run the "FixAudioAnimations" macro once at the end to configure all slides.
                log "Slide " & sIndex & ": Audio added. Pending global VBA fix."
                
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
                log "Audio Duration: " & audioLenSec & "s. Setting advance time to 0 for auto-advance."
                tell slide show transition of currentSlide
                     set advance on time to true
                     set advance time to 0
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
        
        -- STRATEGY V15: SKIPPING XML PATCH
        -- We rely on AppleScript to set animation order and autoplay.
        log "Skipping XML Batch Patch. Relying on AppleScript settings."
        
        -- STRATEGY V28: Robust XML Patcher (Reordering Support)
        -- We will clean up the PPTX using an external Node script that:
        -- 1. Parses the Slide XML
        -- 2. Finds the audio animation node
        -- 3. MOVES it to the start of the Main Sequence (First to play)
        -- 4. Sets trigger to "afterPrev" (Automatic)
        
        log "Applying XML batch patch (Reordering Audio)..."
        
        -- Use Desktop for temp file to guarantee HFS/POSIX path validity and access
        set desktopHFS to path to desktop as text
        set tempHFSPath to desktopHFS & "_temp_patch_work.pptx"
        set tempPptPath to POSIX path of tempHFSPath
        
        log "Saving temp PPTX to HFS path: " & tempHFSPath
        
        -- Cleanup: Close the temp file if it's already open (from a previous failed run)
        try
            close presentation "_temp_patch_work.pptx" saving no
        end try
        
        -- Cleanup existing temp file on disk
        try
            do shell script "rm -f " & quoted form of tempPptPath
        end try
        
        -- Save presentation to temp location so we can patch it
        try
            save pptPres in tempHFSPath
        on error errMsg
            log "Error saving temp PPTX: " & errMsg
            error "Failed to save temp PPTX: " & errMsg
        end try
        
        close pptPres saving no
        
        -- Verify file exists
        try
            do shell script "ls -l " & quoted form of tempPptPath
        on error
            error "Failed to save temp PPTX for patching at: " & tempPptPath
        end try
        
        -- Run XML patcher
        set nodeCmd to "/usr/local/bin/node"
        if not (do shell script "test -f " & quoted form of nodeCmd & " && echo 'yes' || echo 'no'") is "yes" then
            set nodeCmd to "/opt/homebrew/bin/node"
        end if
        
        set patchScript to scriptDir & "/patch-presentation.cjs"
        -- Pass 'all' to process all slides
        set patchCmd to quoted form of nodeCmd & " " & quoted form of patchScript & " " & quoted form of tempPptPath & " all"
        
        try
            do shell script patchCmd
            log "XML patching completed successfully"
        on error patchErr
            log "XML patch warning: " & patchErr
        end try
        
        -- Reopen patched presentation to continue export
        -- Use HFS path here as well for reliability. 
        -- NOTE: 'open' command does not always return the object in all versions.
        open tempHFSPath
        set pptPres to active presentation
        
        -- Writing directly to external folders often fails silently due to sandbox restrictions.
        -- We export to the temp folder, then move it to the destination.
        
        -- Use .mov extension for temp file as we are using 'save as movie'
        set sandboxOutputPOSIX to sandboxTmpPOSIX & "/temp_video.mov"
        -- Construct HFS path for sandbox output
        tell application "System Events" to set startupDisk to name of startup disk
        set sandboxOutputHFS to startupDisk & (do shell script "echo " & quoted form of sandboxOutputPOSIX & " | sed 's/\\//:/g'")
        
        log "Strategy V14: Using 'run script save as movie' for export."
        
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
