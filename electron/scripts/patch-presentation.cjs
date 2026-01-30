const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// Usage: node patch-presentation.js <path-to-pptx> <slide-index>
// Slide index 1-based (default 1)

const filePath = process.argv[2];
const slideIndex = process.argv[3] || 1;

if (!filePath) {
    console.error("Usage: node patch-presentation.js <path-to-pptx> [slide-index]");
    process.exit(1);
}

console.log(`Patching: ${filePath} (Slide ${slideIndex})`);

try {
    const zip = new AdmZip(filePath);
    let modifiedAny = false;

    // Helper function to apply XML patch
    function patchXML(xml) {
        let original = xml;

        // 1. Fix Trigger Type: "On Click" -> "After Previous"
        // The correct attribute value in OpenXML is "afterEffect" (NOT "withPrev")
        // "clickEffect" is the default for on-click animations.
        if (xml.includes('nodeType="clickEffect"')) {
            // Replace ALL instances to ensure audio is caught
            xml = xml.replace(/nodeType="clickEffect"/g, 'nodeType="afterEffect"');
        }

        // 2. Fix Condition: "onClick" -> "begin"
        // This ensures the animation starts automatically when the slide/timeline begins
        // removing the requirement for a mouse click event.
        if (xml.includes('evt="onClick"')) {
            xml = xml.replace(/evt="onClick"/g, 'evt="begin"');
        }

        // 3. Ensure no delay? (Optional, regex replacement for delay="x" is hard safely)

        return { modified: xml !== original, xml };
    }

    if (slideIndex === 'all') {
        const entries = zip.getEntries();
        entries.forEach(entry => {
            // Match ppt/slides/slideX.xml
            if (entry.entryName.match(/^ppt\/slides\/slide\d+\.xml$/)) {
                console.log(`Processing ${entry.entryName}...`);
                const xml = entry.getData().toString('utf8');
                const result = patchXML(xml);
                if (result.modified) {
                    console.log(`  -> Patched audio trigger`);
                    zip.updateFile(entry.entryName, Buffer.from(result.xml, 'utf8'));
                    modifiedAny = true;
                }
            }
        });
    } else {
        // Single slide mode
        const slideEntryName = `ppt/slides/slide${slideIndex}.xml`;
        const slideEntry = zip.getEntry(slideEntryName);
        if (slideEntry) {
            const xml = slideEntry.getData().toString('utf8');
            const result = patchXML(xml);
            if (result.modified) {
                console.log(`  -> Patched audio trigger`);
                zip.updateFile(slideEntryName, Buffer.from(result.xml, 'utf8'));
                modifiedAny = true;
            } else {
                console.log(`  -> No patch needed (already correct?)`);
            }
        } else {
            console.error(`Slide not found: ${slideEntryName}`);
        }
    }

    if (modifiedAny) {
        zip.writeZip(filePath);
        console.log("SUCCESS: PPTX patched successfully.");
    } else {
        console.log("WARNING: No patches needed or applied.");
    }

} catch (e) {
    console.error("Error patching PPTX:", e);
    process.exit(1);
}
