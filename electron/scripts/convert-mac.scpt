on run argv
    set inputPath to item 1 of argv
    set outputDir to item 2 of argv
    
    -- Ensure output directory exists (handled by Posix, but good sanity check)
    do shell script "mkdir -p " & quoted form of outputDir
    
    tell application "Microsoft PowerPoint"
        activate
        -- Open the presentation without showing window if possible (Mac PPT usually shows it)
        open inputPath
        set activePres to active presentation
        
        set slideCount to count of slides of activePres
        set slidesData to {}
        
        repeat with i from 1 to slideCount
            set currentSlide to slide i of activePres
            set imageName to "slide_" & i & ".png"
            set imagePath to outputDir & "/" & imageName
            
            -- Export Slide as Image
            -- Note: AppleScript export syntax varies by version, usually: 
            save currentSlide as PNG to imagePath with height 1080 width 1920
            
            -- Extract Notes
            set notesText to ""
            try
                -- Notes page is technically a separate view
                tell notes page of currentSlide
                    -- Usually shape 2 is the text body placeholder on the notes page
                    set notesText to content of text range of text frame of shape 2
                end tell
            on error
                set notesText to ""
            end try
            
            -- Build JSON object string manually (AppleScript is bad at JSON)
            -- We will construct a simple list string and handle JSON in Node or just print strictly
            -- For simplicity in this script, we output one line per slide for Node to parse
            
            -- Sanitize notes for single line output (basic)
            set cleanNotes to notesText
            
            log "SLIDE_DATA|" & i & "|" & imageName & "|" & cleanNotes
            
        end repeat
        
        close activePres saving no
    end tell
end run
