const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const pptxPath = path.resolve('test_data/test1/test1.pptx');
const tmpDir = path.resolve('temp_pptx_inspect_layout');

if (fs.existsSync(tmpDir)) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
}
fs.mkdirSync(tmpDir);

try {
    execSync(`unzip -q "${pptxPath}" -d "${tmpDir}"`);

    const presXmlPath = path.join(tmpDir, 'ppt', 'presentation.xml');
    if (fs.existsSync(presXmlPath)) {
        console.log(fs.readFileSync(presXmlPath, 'utf8'));
    } else {
        console.log('No ppt/presentation.xml found.');
    }

} catch (err) {
    console.error('Error:', err.message);
}
// valid for manual inspection, won't delete immediately
