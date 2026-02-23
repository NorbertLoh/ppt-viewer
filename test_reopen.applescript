tell application "Microsoft PowerPoint"
    set pres to active presentation
    set presPath to full name of pres
    
    -- Getting the current slide index
    set theSelection to selection of document window 1
    set selectedSlideRange to slide range of theSelection
    set currentSlideIndex to slide index of slide 1 of selectedSlideRange
    
    log "Current Slide Index: " & currentSlideIndex
    log "Presentation Path: " & presPath
    
    close pres saving no
    
    open presPath
    
    set newPres to active presentation
    
    select slide currentSlideIndex of newPres
    
    log "Successfully navigated back to slide " & currentSlideIndex
end tell
