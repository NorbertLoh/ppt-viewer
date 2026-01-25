on run argv
    if count of argv is less than 2 then
        error "Usage: convert-mac.applescript <inputPath> <outputDir>"
    end if
    
    set inputPath to item 1 of argv
    set outputDir to item 2 of argv
    
    -- Ensure output directory exists (handled by TS usually, but good to be safe)
    -- We assume standard POSIX paths
    
    tell application "Microsoft PowerPoint"
        activate
        
        -- Open the presentation
        try
            set pptPres to open (POSIX file inputPath)
        on error errMsg
            error "Failed to open presentation: " & errMsg
        end try
        
        set slideCount to count of slides of pptPres
        set slidesData to {}
        
        repeat with i from 1 to slideCount
            set currentSlide to slide i of pptPres
            
            -- Image Path
            set imageName to "slide_" & i & ".png"
            set imagePath to outputDir & "/" & imageName
            
            -- Export Slide as PNG
            -- PowerPoint Mac 'save' command with 'as save as PNG' saves the whole generic presentation format
            -- To save a specific slide, we often use 'export' or specific save commands.
            -- Actually, Mac PPT has a 'save slide' command or similar?
            -- It seems 'save' on the slide object works best.
            
            try
                 save currentSlide in (POSIX file imagePath) as save as PNG
            on error errMsg
                 -- Sometimes creating the file fails if folder issues, trying generic save
            end try
            
            -- Extract Notes
            set notesText to ""
            try
                -- Notes Page -> Shape 2 (body) is the standard pattern
                -- We iterate shapes to find the body placeholder if needed, but standard is fixed.
                set notesPage to notes page of currentSlide
                -- Shape 2 is typically the text body in default layout
                set notesText to content of text range of text frame of shape 2 of notesPage
            on error
                -- Ignore if no notes or different layout
            end try
            
            -- Build JSON Object String (manual because AppleScript JSON is hard)
            -- We need to escape quotes in notes
            set escapedNotes to my findAndReplace(notesText, "\"", "\\\"")
            set escapedNotes to my findAndReplace(escapedNotes, "
", "\\n") -- Handle newlines
            
            set slideJson to "{ \"index\": " & i & ", \"image\": \"" & imageName & "\", \"notes\": \"" & escapedNotes & "\" }"
            set end of slidesData to slideJson
            
        end repeat
        
        -- Close PPT
        close pptPres saving no
        -- quit -- We might not want to quit if user has other things open, but consistent with Win script
        quit
        
    end tell
    
    -- Write Manifest JSON
    -- Join list with commas
    set jsonString to "[" & my joinList(slidesData, ",") & "]"
    set manifestPath to outputDir & "/manifest.json"
    
    try
        set jsonFile to open for access (POSIX file manifestPath) with write permission
        set eof jsonFile to 0
        write jsonString to jsonFile as «class utf8»
        close access jsonFile
    on error errMsg
        try
            close access jsonFile
        end try
        error "Failed to write manifest: " & errMsg
    end try
    
    return manifestPath
end run

-- Helper: Find and Replace
on findAndReplace(txt, findStr, replaceStr)
    set {oldTID, AppleScript's text item delimiters} to {AppleScript's text item delimiters, findStr}
    set txtItems to text items of txt
    set AppleScript's text item delimiters to replaceStr
    set newTxt to txtItems as string
    set AppleScript's text item delimiters to oldTID
    return newTxt
end findAndReplace

-- Helper: Join List
on joinList(theList, delimiter)
    set {oldTID, AppleScript's text item delimiters} to {AppleScript's text item delimiters, delimiter}
    set theString to theList as string
    set AppleScript's text item delimiters to oldTID
    return theString
end joinList
