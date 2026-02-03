function run(argv) {
    if (argv.length < 2) {
        throw new Error("Usage: save-all-notes-mac.js <inputPath> <notesJsonPath>");
    }

    var inputPath = argv[0];
    var jsonPath = argv[1];

    var app = Application('Microsoft PowerPoint');
    // app.includeStandardAdditions = true; // Use currentApp for file IO

    var currentApp = Application.currentApplication();
    currentApp.includeStandardAdditions = true;

    // Read JSON content using currentApp (osascript) to avoid sandbox issues
    var jsonContent = currentApp.read(Path(jsonPath));
    var updates = JSON.parse(jsonContent);

    // app.activate(); // REMOVED: Do not steal focus

    var pptPres = null;
    var wasAlreadyOpen = false;

    // Check if presentation is already open
    var openPresCount = app.presentations.length;
    for (var i = 0; i < openPresCount; i++) {
        var p = app.presentations[i];
        // Check path (Mac paths might vary slightly, checking name or full path)
        try {
            if (p.fullName() === inputPath || p.path() === inputPath) {
                pptPres = p;
                wasAlreadyOpen = true;
                break;
            }
        } catch (e) { }
    }

    if (!pptPres) {
        // Open Presentation (hidden if possible, though JXA 'open' usually shows window)
        // We can minimize it? Or just let it run in bg.
        pptPres = app.open(Path(inputPath));
        // pptPres.window.minimized = true; // Optional: try to minimize
    }

    // Iterate updates
    updates.forEach(function (item) {
        var idx = item.index;
        var newNotes = item.notes;

        try {
            // Slide indices are 1-based in PPT DOM usually
            // Access by index seems to be 0-based in JXA arrays, but PPT collections can be 1-based.
            // Let's rely on standard collection access.
            // If we generated indices as 1-based, we might need [idx-1].
            // BUT, if we are unsure, we can try to find by index property if available.
            // Let's assume input 'idx' is 1-based (slide number).

            var slide = pptPres.slides[idx - 1];

            // Notes Page
            var notesPage = slide.notesPage;
            // Shape 2 is body text placeholder usually
            var notesShape = notesPage.shapes[1]; // 2nd shape

            // Set Text
            notesShape.textFrame.textRange.content = newNotes;
        } catch (e) {
            console.log("Error updating slide " + idx + ": " + e.message);
        }
    });

    // Save
    pptPres.save();

    // Close ONLY if we opened it
    if (!wasAlreadyOpen) {
        pptPres.close();
    }

    // app.quit(); // REMOVED: Keep app running
}
