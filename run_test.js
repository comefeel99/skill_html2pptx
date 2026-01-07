const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const html2pptx = require('./pptx/scripts/html2pptx');
const { comparePptx } = require('./compare_pptx'); // Import function
const { generateReport } = require('./generate_report'); // Import function

async function runTest() {
    const startTime = Date.now();
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
    let generatedCount = 0;

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
            generatedCount++;
        } catch (err) {
            console.error(`Error converting ${file}:`, err);
        }
    }

    await pres.writeFile({ fileName: outputPath });
    const outputSize = fs.statSync(outputPath).size;
    console.log(`Saved output to ${outputPath} (${outputSize} bytes)`);

    // --- Comparison Step ---
    let comparisonResult = null;
    if (fs.existsSync(referencePath)) {
        console.log(`\n=== Comparing with Reference: ${referencePath} ===`);
        try {
            comparisonResult = comparePptx(referencePath, outputPath);

            // Console output for user visibility
            console.log(`Overall Similarity: ${comparisonResult.similarity}%`);

            // Show slide details
            const slideKeys = Object.keys(comparisonResult.slides).sort();
            console.log(`Slides evaluated: ${slideKeys.length}`);
            if (slideKeys.length > 0) {
                console.log('Slide Scores:', slideKeys.map(k => `${k}: ${comparisonResult.slides[k].similarity}%`).join(', '));
            }

        } catch (err) {
            console.error('Comparison execution failed:', err);
            comparisonResult = { error: err.message };
        }
    } else {
        console.log(`\nNo reference file found at ${referencePath}. Skipping comparison.`);
    }

    const durationMs = Date.now() - startTime;

    // --- Logging Step ---
    const logDir = path.resolve(__dirname, 'test_result');
    if (!fs.existsSync(logDir)) fs.mkdirSync(logDir, { recursive: true });

    const logEntry = {
        timestamp: new Date().toISOString(),
        durationMs: durationMs,
        testDir: dirName,
        generatedFile: path.basename(outputPath),
        generatedSize: outputSize,
        generatedSlides: generatedCount,
        inputFiles: files.length,
        comparison: comparisonResult
    };

    const logPath = path.join(logDir, 'history.jsonl');
    fs.appendFileSync(logPath, JSON.stringify(logEntry) + '\n');
    console.log(`\nTest result logged to ${logPath}`);

    // --- Report Generation Step ---
    generateReport();
}

runTest().catch(console.error);
