const pptxgen = require('pptxgenjs');
const path = require('path');
const html2pptx = require('./pptx/scripts/html2pptx');
const fs = require('fs');

async function convertTestData() {
    const testDataDir = path.resolve(__dirname, 'test_data/test1');
    console.log(`Starting conversion of 9 pages from ${testDataDir} to separate PPTX files...`);

    // We will read the HTML files to find their dimensions
    // But html2pptx uses playwright to get computed styles, so we can't easily regex it.
    // We'll rely on a reasonable default scaling or just hardcode for known pages if needed.
    // Better: create a new pptx instance for each page.
    // We can infer dimensions or pass a "flexible" validation expectation?
    // Since we disabled validation errors in html2pptx.js, we can just proceed.
    // But we WANT correct layout.
    // We can use a helper to peek dimensions or just try to match commonly known sizes.
    // Page 1-7: 1280x728.
    // Page 8: 1280x872.
    // Page 9: 1280x1086.

    // We will hardcode dimensions based on our inspection to ensure best results
    const dimensions = {
        'page_1.html': { w: 1280, h: 728 },
        'page_2.html': { w: 1280, h: 728 }, // Assuming same
        'page_3.html': { w: 1280, h: 728 },
        'page_4.html': { w: 1280, h: 728 },
        'page_5.html': { w: 1280, h: 728 },
        'page_6.html': { w: 1280, h: 728 },
        'page_7.html': { w: 1280, h: 728 },
        'page_8.html': { w: 1280, h: 872 },
        'page_9.html': { w: 1280, h: 1086 }
    };

    for (let i = 1; i <= 9; i++) {
        const filename = `page_${i}.html`;
        const filePath = path.join(testDataDir, filename);
        if (!fs.existsSync(filePath)) {
            console.log(`Skipping ${filename} (not found)`);
            continue;
        }

        console.log(`Processing ${filename}...`);

        try {
            const pptx = new pptxgen();

            const dims = dimensions[filename] || { w: 1280, h: 728 };
            const wInch = dims.w / 96;
            const hInch = dims.h / 96;

            const layoutName = `LAYOUT_${i}`;
            pptx.defineLayout({ name: layoutName, width: wInch, height: hInch });
            pptx.layout = layoutName;

            await html2pptx(filePath, pptx);

            const outputFile = `page_${i}.pptx`;
            await pptx.writeFile({ fileName: outputFile });
            console.log(`Saved ${outputFile}`);
        } catch (err) {
            console.error(`Error converting ${filename}:`, err);
        }
    }

    console.log("All conversions attempted.");
}

convertTestData().catch(err => {
    console.error("Fatal error:", err);
});
