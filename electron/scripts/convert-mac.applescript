on run argv
    if count of argv is less than 2 then
        error "Usage: convert-mac.applescript <inputPath> <outputDir>"
    end if
    
    set inputPath to item 1 of argv
    set outputDir to item 2 of argv
    
    -- 1. Ensure output directory exists (Parent of where we want the slides folder)
    do shell script "mkdir -p " & quoted form of outputDir
    
    -- 2. Define the target "slides" folder path
    -- We want: outputDir/slides
    -- First, get the HFS path of outputDir (which exists)
    set outputHFS to ""
    tell application "System Events"
        if exists disk item outputDir then
            set outputHFS to path of disk item outputDir as string
        else
            error "Output dir does not exist: " & outputDir
        end if
    end tell
    
    -- outputHFS ends with ":" if it is a folder.
    -- Append "slides" to create the target folder path
    set destinationHFS to outputHFS & "slides"
    
    -- 3. Clean up previous attempt
    do shell script "rm -rf " & quoted form of (outputDir & "/slides")
    
    tell application "Microsoft PowerPoint"
        activate
        
        -- Open the presentation
        try
            set pptPres to open (POSIX file inputPath)
        on error errMsg
            error "Failed to open presentation: " & errMsg
        end try
        
        -- EXPORT AS PNG (Bulk)
        -- We use the 'file' specifier with the string path.
        -- This tells AppleScript it's a file path, even if it doesn't exist yet.
        try
             save pptPres in file destinationHFS as save as PNG
        on error errMsg
             error "Failed to export slides (Bulk Save): " & errMsg & " (Dest: " & destinationHFS & ")"
        end try
        
        set slideCount to count of slides of pptPres
        set slidesData to {}
        
        repeat with i from 1 to slideCount
            set currentSlide to slide i of pptPres
            
            -- Image Path Logic:
            -- Mac PowerPoint Bulk Export creates "Slide1.PNG", "Slide2.PNG" ...
            -- INSIDE the target folder we specified ("slides").
            set imageName to "slides/Slide" & i & ".PNG"
            
            -- Extract Notes
            set notesText to ""
            try
                set notesPage to notes page of currentSlide
                set notesText to content of text range of text frame of shape 2 of notesPage
            on error
                -- Ignore
            end try
            
            -- Escape JSON
            set escapedNotes to my findAndReplace(notesText, "\"", "\\\"")
            set escapedNotes to my findAndReplace(escapedNotes, "
", "\\n")
            set escapedNotes to my findAndReplace(escapedNotes, return, "\\n")
            
            set slideJson to "{ \"index\": " & i & ", \"image\": \"" & imageName & "\", \"notes\": \"" & escapedNotes & "\" }"
            set end of slidesData to slideJson
            
        end repeat
        
        -- Close PPT
        close pptPres saving no
        
    end tell
    
    -- Write Manifest JSON
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
