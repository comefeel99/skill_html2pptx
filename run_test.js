const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');
const html2pptx = require('./pptx/scripts/html2pptx');

async function runTest() {
    const testDirArg = process.argv[2];
    if (!testDirArg) {
        console.error('Usage: node run_test.js <test_directory_path>');
        process.exit(1);
    }

    const testDir = path.resolve(testDirArg);
    if (!fs.existsSync(testDir)) {
        console.error(`Directory not found: ${testDir}`);
        process.exit(1);
    }

    const dirName = path.basename(testDir);
    const outputPath = path.resolve(__dirname, `${dirName}_output.pptx`);
    const referencePath = path.join(testDir, `${dirName}.pptx`);

    // --- Generation Step ---
    console.log(`\n=== Generating PPTX from ${testDir} ===`);

    const pres = new PptxGenJS();
    pres.defineLayout({ name: 'TEST_LAYOUT', width: 13.3333, height: 7.5 });
    pres.layout = 'TEST_LAYOUT';

    const files = fs.readdirSync(testDir)
        .filter(file => file.startsWith('page_') && file.endsWith('.html'))
        .sort((a, b) => {
            const numA = parseInt(a.match(/page_(\d+)\.html/)[1]);
            const numB = parseInt(b.match(/page_(\d+)\.html/)[1]);
            return numA - numB;
        });

    console.log(`Found ${files.length} HTML files.`);

    for (const file of files) {
        const filePath = path.join(testDir, file);
        const stats = fs.statSync(filePath);

        // Skip empty or very small files (less than ~100 bytes is likely just tags)
        if (stats.size < 100) {
            console.log(`Skipping ${file} (Size: ${stats.size} bytes - considered empty)...`);
            continue;
        }

        const pageNum = parseInt(file.match(/page_(\d+)\.html/)[1]);
        console.log(`Processing ${file} into slide ${pageNum}...`);

        try {
            const slide = pres.addSlide();
            await html2pptx(filePath, pres, {
                slide: slide,
                masterOptions: { margin: 0 }
            });
        } catch (err) {
            console.error(`Error converting ${file}:`, err);
        }
    }

    await pres.writeFile({ fileName: outputPath });
    console.log(`Saved output to ${outputPath}`);

    // --- Comparison Step ---
    if (fs.existsSync(referencePath)) {
        console.log(`\n=== Comparing with Reference: ${referencePath} ===`);
        try {
            // Run compare_pptx.js
            // Usage: node compare_pptx.js <file1> <file2>
            const compareScript = path.join(__dirname, 'compare_pptx.js');
            // Execute simply and inherit stdio to show output directly
            execSync(`node "${compareScript}" "${referencePath}" "${outputPath}"`, { stdio: 'inherit' });
        } catch (err) {
            console.error('Comparison execution failed:', err);
        }
    } else {
        console.log(`\nNo reference file found at ${referencePath}. Skipping comparison.`);
    }
}

runTest().catch(console.error);
