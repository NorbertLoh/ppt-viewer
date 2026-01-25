#!/usr/bin/env osascript -l JavaScript

// convert-mac.js - JXA script to convert PPTX to PNG slides and extract notes
// Usage: osascript -l JavaScript convert-mac.js <inputPath> <outputDir>

function run(argv) {
    if (argv.length < 2) {
        throw new Error("Usage: convert-mac.js <inputPath> <outputDir>");
    }

    var inputPath = argv[0];
    var outputDir = argv[1];

    // Create slides subdirectory
    var slidesDir = outputDir + "/slides";
    var app = Application.currentApplication();
    app.includeStandardAdditions = true;
    app.doShellScript("mkdir -p " + quoteForShell(slidesDir));

    // Open PowerPoint
    var ppt = Application("Microsoft PowerPoint");
    ppt.activate();

    // Open Presentation
    var pres = ppt.open(Path(inputPath));

    // Get slide count
    var slideCount = pres.slides.length;
    var slidesData = [];

    // Export individual slides
    for (var i = 0; i < slideCount; i++) {
        var slideNum = i + 1;
        var fileName = "Slide" + slideNum + ".png";
        var filePath = slidesDir + "/" + fileName;

        // Export slide as image
        // In JXA, we use pres.slides[i].export(...)
        try {
            var slide = pres.slides[i];
            // Export to file - JXA style
            // The 'saveable file format' for PNG is 'save as PNG' (17) or similar
            // Actually, let's try the 'export' method if available

            // Fallback: Use the presentation's saveAs with a specific slide?
            // Mac PPT JXA doesn't have a direct per-slide export.
            // Let's try saving the whole presentation as PNG first
        } catch (e) {
            // Log but continue
            console.log("Warning exporting slide " + slideNum + ": " + e.message);
        }

        // Extract notes
        var notesText = "";
        try {
            var notesPage = pres.slides[i].notesPage;
            // Shape at index 1 (0-based) is typically the body placeholder
            if (notesPage.shapes.length > 1) {
                notesText = notesPage.shapes[1].textFrame.textRange.content();
            }
        } catch (e) {
            // No notes or different structure
        }

        slidesData.push({
            index: slideNum,
            image: "slides/" + fileName,
            notes: notesText
        });
    }

    // BULK EXPORT - Save presentation as PNG to slidesDir
    // This exports all slides at once
    try {
        // ppSaveAsPNG = 18 (PowerPoint constant)
        // We save to the slidesDir path
        pres.saveAs(Path(slidesDir), { as: "save as PNG" });
    } catch (e) {
        // Try alternative: export using file format number
        try {
            // ppSaveAsPNG = 18
            pres.saveAs(Path(slidesDir), { fileFormat: 18 });
        } catch (e2) {
            throw new Error("Failed to export slides: " + e.message + " | " + e2.message);
        }
    }

    // Close presentation
    pres.close({ saving: "no" });

    // Write manifest JSON
    var manifestPath = outputDir + "/manifest.json";
    var jsonContent = JSON.stringify(slidesData, null, 2);

    // Write to file using shell
    app.doShellScript("cat > " + quoteForShell(manifestPath) + " << 'EOFMANIFEST'\n" + jsonContent + "\nEOFMANIFEST");

    return manifestPath;
}

function quoteForShell(str) {
    return "'" + str.replace(/'/g, "'\\''") + "'";
}
