tell application "Microsoft PowerPoint"
    set myName to name of active presentation
    set ppath to "Macintosh HD:private:tmp:test_ppt_export:"
    do shell script "mkdir -p /tmp/test_ppt_export"
    save active presentation in ppath as save as PNG
    do shell script "find /tmp/test_ppt_export -type f"
end tell
