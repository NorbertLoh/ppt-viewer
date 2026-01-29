function run(argv) {
    if (argv.length < 2) {
        throw new Error("Usage: save-all-notes-mac.js <inputPath> <notesJsonPath>");
    }

    var inputPath = argv[0];
    var jsonPath = argv[1];

    var app = Application('Microsoft PowerPoint');
    app.includeStandardAdditions = true;

    // Read JSON content
    var jsonContent = app.read(Path(jsonPath));
    var updates = JSON.parse(jsonContent);

    app.activate();

    // Open Presentation
    // JXA open command
    var pptPres = app.open(Path(inputPath));

    // Iterate updates
    updates.forEach(function (item) {
        var idx = item.index;
        var newNotes = item.notes;

        // Slide indices are 1-based in PPT, check if JSON provided 1-based (usually is from our code)
        var slide = pptPres.slides[idx - 1]; // JXA arrays are 0-based? 
        // Wait, 'pptPres.slides' is a collection.
        // Collection access in JXA can be by index (0-based) or by name?
        // Usually JXA collections are 0-based access for `[i]`.
        // BUT PowerPoint DOM might be 1-based if using .item()?
        // Let's safe-check. If idx is 1 (Slide 1), we want 0-th element probably.

        // Actually, reliable way: `pptPres.slides[idx - 1]`

        try {
            // Notes Page
            var notesPage = slide.notesPage;
            // Shape 2 is body
            // In JXA, shapes is a collection
            // `notesPage.shapes[1]` (0-based index 1 = 2nd item)?
            // Or `notesPage.shapes.item(2)`?

            // Let's assume standard JXA array access
            var notesShape = notesPage.shapes[1]; // 2nd shape

            // Set Text
            // notesShape.textFrame.textRange.content = newNotes
            notesShape.textFrame.textRange.content = newNotes;
        } catch (e) {
            console.log("Error updating slide " + idx + ": " + e.message);
        }
    });

    // Save and Close
    pptPres.save();
    pptPres.close();
    app.quit();
}
