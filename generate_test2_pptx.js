const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const html2pptx = require('./pptx/scripts/html2pptx');

async function generateTest2Pptx() {
    const pres = new PptxGenJS();

    // Define 16:9 layout matching 1280x720 aspect ratio
    // 13.333 inches * 72 dpi = 960 pixels. 7.5 inches * 72 dpi = 540 pixels. 
    // 960/540 = 1.777. 1280/720 = 1.777. Correct.
    pres.defineLayout({ name: 'TEST2_LAYOUT', width: 13.3333, height: 7.5 });
    pres.layout = 'TEST2_LAYOUT';

    const testDataDir = path.join(__dirname, 'test_data/test2');
    const files = fs.readdirSync(testDataDir)
        .filter(file => file.startsWith('page_') && file.endsWith('.html'))
        .sort((a, b) => {
            const numA = parseInt(a.match(/page_(\d+)\.html/)[1]);
            const numB = parseInt(b.match(/page_(\d+)\.html/)[1]);
            return numA - numB;
        });

    console.log(`Found ${files.length} HTML files in ${testDataDir}`);
    console.log('Starting conversion of test2 pages to test2_output.pptx...');

    for (const file of files) {
        const pageNum = parseInt(file.match(/page_(\d+)\.html/)[1]);

        // Skip page 13 as per plan (it was empty/small in list_dir)
        if (pageNum === 13) {
            console.log(`Skipping ${file} (Page 13)...`);
            continue;
        }

        const filePath = path.join(testDataDir, file);
        // const html = fs.readFileSync(filePath, 'utf-8'); // Removed

        console.log(`Processing ${file} into slide ${pageNum}...`);

        try {
            const slide = pres.addSlide();
            await html2pptx(filePath, pres, {
                slide: slide,
                masterOptions: {
                    margin: 0 // No margin for full slides
                }
            });
        } catch (err) {
            console.error(`Error converting ${file}:`, err);
        }
    }

    const outputPath = path.join(__dirname, 'test2_output.pptx');
    await pres.writeFile({ fileName: outputPath });
    console.log(`\nSuccessfully saved Test 2 PPTX to ${outputPath}`);
}

generateTest2Pptx().catch(console.error);
