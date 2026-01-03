const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const pptxPath = path.resolve('test_data/test1/test1.pptx');
const tmpDir = path.resolve('temp_pptx_inspect');

if (fs.existsSync(tmpDir)) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
}
fs.mkdirSync(tmpDir);

try {
    // PPTX is a zip file. Unzip it.
    execSync(`unzip -q "${pptxPath}" -d "${tmpDir}"`);

    const slidesDir = path.join(tmpDir, 'ppt', 'slides');
    if (fs.existsSync(slidesDir)) {
        const files = fs.readdirSync(slidesDir);
        const slideFiles = files.filter(f => f.startsWith('slide') && f.endsWith('.xml'));
        console.log(`Slide count in test1.pptx: ${slideFiles.length}`);
        console.log('Slides:', slideFiles.sort());
    } else {
        console.log('No ppt/slides directory found.');
    }

} catch (err) {
    console.error('Error inspecting PPTX:', err.message);
} finally {
    // Cleanup
    if (fs.existsSync(tmpDir)) {
        fs.rmSync(tmpDir, { recursive: true, force: true });
    }
}
