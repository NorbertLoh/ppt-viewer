tell application "Microsoft PowerPoint"
    launch
    set newPres to make new presentation
    set s1 to make new slide at end of newPres with properties {layout:slide layout blank}
    make new shape at s1 with properties {auto shape type:shape rectangle, text frame:{text range:{content:"Slide A"}}}
    set s2 to make new slide at end of newPres with properties {layout:slide layout blank}
    make new shape at s2 with properties {auto shape type:shape rectangle, text frame:{text range:{content:"Slide B"}}}
    set s3 to make new slide at end of newPres with properties {layout:slide layout blank}
    make new shape at s3 with properties {auto shape type:shape rectangle, text frame:{text range:{content:"Slide C"}}}
    
    -- Rename them so we know their original creation
    set name of s1 to "Orig1"
    set name of s2 to "Orig2"
    set name of s3 to "Orig3"
    
    -- Move Slide 3 to position 1
    -- Well, AppleScript to move slides in PPT Mac is tricky. Let's try cut/paste or just `move slide 3 to to index 1` if it exists.
end tell
