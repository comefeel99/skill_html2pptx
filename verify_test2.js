const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Assuming inspect_pptx.js logic or similar
// We will just unzip to a temp dir and count slides and check media

const pptxPath = path.join(__dirname, 'test2_output.pptx');
const tempDir = path.join(__dirname, 'temp_inspect_test2');

if (fs.existsSync(tempDir)) {
    fs.rmSync(tempDir, { recursive: true, force: true });
}
fs.mkdirSync(tempDir);

try {
    // Unzip
    execSync(`unzip -q ${pptxPath} -d ${tempDir}`);

    // Count slides
    const slidesDir = path.join(tempDir, 'ppt/slides');
    const slides = fs.readdirSync(slidesDir).filter(f => f.startsWith('slide') && f.endsWith('.xml'));
    console.log(`\nSlide Content Analysis:`);
    console.log(`Total Slides: ${slides.length} (Expected 12)`);

    // Check Media
    const mediaDir = path.join(tempDir, 'ppt/media');
    let mediaCount = 0;
    if (fs.existsSync(mediaDir)) {
        mediaCount = fs.readdirSync(mediaDir).length;
    }
    console.log(`Total Media Files: ${mediaCount}`);

    // Analyze each slide size to detect failures
    slides.sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml/)[1]);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)[1]);
        return numA - numB;
    });

    for (const s of slides) {
        const sPath = path.join(slidesDir, s);
        const content = fs.readFileSync(sPath, 'utf-8');
        const size = content.length;
        const hasImage = content.includes('<p:pic') || content.includes('<a:blip');
        const hasText = content.includes('<a:t>');

        console.log(`  ${s}: Size=${size} bytes, HasImage=${hasImage}, HasText=${hasText}`);

        if (size < 2000) {
            console.warn(`    WARNING: Slide ${s} seems very empty! Possible conversion failure.`);
        }
    }

} catch (err) {
    console.error("Verification failed:", err);
} finally {
    // Cleanup
    if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
    }
}
