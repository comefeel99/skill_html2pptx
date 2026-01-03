const pptxgen = require('pptxgenjs');
const path = require('path');
const html2pptx = require('./pptx/scripts/html2pptx');
const fs = require('fs');

async function generateCombinedPptx() {
    const testDataDir = path.resolve(__dirname, 'test_data/test1');
    const outputFile = 'test_output_combined.pptx';

    console.log(`Starting combined conversion of 9 pages to ${outputFile}...`);

    // Create a SINGLE presentation instance
    const pres = new pptxgen();

    // Set the layout strictly to match test1.pptx (13.333... x 7.5 inches)
    // 13.333333333333334 inches = 1280px / 96dpi
    // 7.5 inches = 720px / 96dpi ? No, 7.5 * 96 = 720 px.
    // page_1.html has min-height 728px. 
    // The inspected xml said cy="6858000" -> 7.5 inches.
    // 7.5 inches * 96 = 720 px.
    // If the content is 728px, it will slightly overflow vertically (8px). This is expected.

    // Define and set the layout
    pres.defineLayout({ name: 'TEST1_LAYOUT', width: 13.33333333, height: 7.5 });
    pres.layout = 'TEST1_LAYOUT';

    for (let i = 1; i <= 9; i++) {
        const filename = `page_${i}.html`;
        const filePath = path.join(testDataDir, filename);

        if (!fs.existsSync(filePath)) {
            console.log(`Skipping ${filename} (not found)`);
            continue;
        }

        console.log(`Processing ${filename} into slide ${i}...`);

        try {
            // Add a new slide to the existing presentation
            // Note: html2pptx normally takes 'pres' to add slide, or we need to manage it.
            // Looking at html2pptx.js usage:
            // await html2pptx(htmlFile, pres, options);
            // It calls pres.addSlide() internally if options.slide is not provided?
            // Let's verify html2pptx.js code.
            // It has: const slide = options.slide || pres.addSlide();
            // So passing 'pres' is correct, it will append a slide.

            await html2pptx(filePath, pres);

        } catch (err) {
            console.error(`Error converting ${filename}:`, err);
        }
    }

    await pres.writeFile({ fileName: outputFile });
    console.log(`Successfully saved combined PPTX to ${outputFile}`);
}

generateCombinedPptx().catch(err => {
    console.error("Fatal error:", err);
});
