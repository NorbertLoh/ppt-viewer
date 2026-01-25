on run argv
    -- Argument Handling
    if count of argv is less than 3 then
        error "Usage: generate-video-mac.applescript <inputPath> <audioDir> <outputPath>"
    end if
    
    set inputPath to item 1 of argv
    set audioDir to item 2 of argv
    set outputPath to item 3 of argv
    
    tell application "Microsoft PowerPoint"
        activate
        
        -- Open the presentation
        -- We use `open` with a POSIX file object
        try
            set pptPres to open (POSIX file inputPath)
        on error errMsg
            error "Failed to open presentation: " & errMsg
        end try
        
        -- Iterate through slides
        set slideCount to count of slides of pptPres
        
        repeat with i from 1 to slideCount
            set currentSlide to slide i of pptPres
            set audioFile to audioDir & "/slide_" & i & ".wav"
            
            -- Check if audio file exists using System Events
            set audioExists to false
            tell application "System Events"
                if exists file audioFile then
                    set audioExists to true
                end if
            end tell
            
            if audioExists then
                -- Add Audio Media Object
                -- Note: coordinates are points. top/left 10 is off-center but visible.
                set mediaShape to make new media object at currentSlide with properties {file name:audioFile, top:10, left:10, width:50, height:50, link to file:false, save with document:true}
                
                -- Configure Animation Settings
                -- We want Play On Entry and Hide While Not Playing
                tell animation settings of mediaShape
                     set animate to true
                     tell play settings
                          set play on entry to true
                          set hide while not playing to true
                     end tell
                end tell
                
                -- Get Audio Duration to set Slide Transition
                set audioLenMs to 0
                try
                    -- length is usually in milliseconds
                    set audioLenMs to length of media format of mediaShape
                end try
                
                set audioLenSec to audioLenMs / 1000
                
                -- Set Slide Transition to Advance on Time
                tell slide show transition of currentSlide
                     set advance on time to true
                     set advance time to audioLenSec
                end tell
                
            else
                -- No Audio: Default Transition
                tell slide show transition of currentSlide
                     if advance on time is false then
                         set advance on time to true
                         set advance time to 3.0 -- Default 3 seconds
                     end if
                end tell
            end if
        end repeat
        
        -- Export Video
        -- 'save as mp4' is the standard enum for recent PPT versions
        try
            save pptPres in (POSIX file outputPath) as save as mp4
        on error errMsg
            error "Failed to export video: " & errMsg
        end try
        
        -- Wait for Export Completion
        -- The `save` command returns immediately for video exports on Mac.
        -- We need to wait for the file to be fully written.
        
        -- Give it a moment to start
        delay 2
        
        set exportComplete to false
        set previousSize to -1
        set stableCount to 0
        
        -- Timeout loop (approx 5 minutes)
        repeat 300 times
            tell application "System Events"
                if exists file outputPath then
                    set currentSize to size of file outputPath
                else
                    set currentSize to -1
                end if
            end tell
            
            if currentSize > 0 then
                if currentSize is equal to previousSize then
                    set stableCount to stableCount + 1
                else
                    set stableCount to 0
                end if
                
                set previousSize to currentSize
            end if
            
            -- If size is stable for 5 seconds, assume done
            if stableCount ge 5 then
                set exportComplete to true
                exit repeat
            end if
            
            delay 1
        end repeat
        
        if not exportComplete then
             error "Video export timed out or failed to detect completion."
        end if
        
        -- Close without saving changes to the PPT text
        close pptPres saving no
        
        -- Quit PowerPoint (matches the behavior of the win script)
        quit
        
    end tell
end run
