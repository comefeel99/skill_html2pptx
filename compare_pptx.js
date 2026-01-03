const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const file1 = process.argv[2] ? path.resolve(process.argv[2]) : path.resolve('test_data/test1/test1.pptx');
const file2 = process.argv[3] ? path.resolve(process.argv[3]) : path.resolve('test_output_combined.pptx');
const outDir = path.resolve('temp_compare');

function unzip(file, dir) {
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    execSync(`unzip -q "${file}" -d "${dir}"`);
}

function countFiles(dir) {
    if (!fs.existsSync(dir)) return 0;
    return fs.readdirSync(dir).length;
}

function getSlideSizes(dir) {
    const presXml = path.join(dir, 'ppt', 'presentation.xml');
    if (!fs.existsSync(presXml)) return 'Unknown';
    const content = fs.readFileSync(presXml, 'utf8');
    const match = content.match(/sldSz cx="(\d+)" cy="(\d+)"/);
    if (match) {
        return `${Math.round(match[1] / 914400)}x${Math.round(match[2] / 914400)} inches`;
    }
    return 'Unknown';
}

function getSlideContentFeatures(dir) {
    const slidesDir = path.join(dir, 'ppt', 'slides');
    if (!fs.existsSync(slidesDir)) return {};

    const results = {};
    fs.readdirSync(slidesDir).forEach(f => {
        if (!f.endsWith('.xml')) return;
        const content = fs.readFileSync(path.join(slidesDir, f), 'utf8');

        // Count occurrences using regex
        const spCount = (content.match(/<p:sp>/g) || []).length;
        const picCount = (content.match(/<p:pic>/g) || []).length;
        const grpSpCount = (content.match(/<p:grpSp>/g) || []).length;
        const txBodyCount = (content.match(/<p:txBody>/g) || []).length;
        const hasBgPr = content.includes('<p:bgPr>');
        const hasBgRef = content.includes('<p:bgRef>');

        results[f] = {
            size: content.length,
            shapes: spCount,
            images: picCount,
            groups: grpSpCount,
            textBlocks: txBodyCount,
            hasBackground: hasBgPr || hasBgRef
        };
    });
    return results;
}

function dumpSlideText(dir, slideName) {
    const slidePath = path.join(dir, 'ppt', 'slides', slideName);
    if (!fs.existsSync(slidePath)) {
        console.log(`Slide ${slideName} not found in ${dir}`);
        return;
    }
    const content = fs.readFileSync(slidePath, 'utf8');
    // Simple regex to grab text inside <a:t>
    const texts = content.match(/<a:t>(.*?)<\/a:t>/g) || [];
    console.log(`--- Text Content for ${slideName} ---`);
    texts.forEach(t => {
        const clean = t.replace(/<\/?a:t>/g, '');
        console.log(`[${clean}]`);
    });
    console.log('-----------------------------------');
}

try {
    console.log(`Comparing:\n A: ${file1}\n B: ${file2}\n`);

    const dir1 = path.join(outDir, 'A');
    const dir2 = path.join(outDir, 'B');

    unzip(file1, dir1);
    unzip(file2, dir2);

    // 1. Dimensions
    console.log(`Dimensions (A): ${getSlideSizes(dir1)}`);
    console.log(`Dimensions (B): ${getSlideSizes(dir2)}`);

    // 2. Media Count
    const media1 = countFiles(path.join(dir1, 'ppt', 'media'));
    const media2 = countFiles(path.join(dir2, 'ppt', 'media'));
    console.log(`Media Files (A): ${media1}`);
    console.log(`Media Files (B): ${media2} (Diff: ${media2 - media1})`);

    // 3. Slide Analysis
    const slides1 = getSlideContentFeatures(dir1);
    const slides2 = getSlideContentFeatures(dir2);

    console.log('\nSlide Content Analysis:');
    Object.keys(slides1).sort().forEach(key => {
        const s1 = slides1[key];
        const s2 = slides2[key];

        if (!s2) {
            console.log(`Slide ${key}: Only in A`);
            return;
        }

        const diffSize = s2.size - s1.size;

        console.log(`Slide ${key}:`);
        console.log(`  Size:     A=${s1.size}, B=${s2.size} (Diff: ${diffSize})`);
        console.log(`  Images:   A=${s1.images}, B=${s2.images}`);
        console.log(`  Shapes:   A=${s1.shapes}, B=${s2.shapes}`);
        console.log(`  Groups:   A=${s1.groups}, B=${s2.groups}`);
        console.log(`  Texts:    A=${s1.textBlocks}, B=${s2.textBlocks}`);
        console.log(`  Backgrnd: A=${s1.hasBackground}, B=${s2.hasBackground}`);
    });

    // 4. Detailed Text Dump (Optional)
    const specificSlide = process.argv[4]; // e.g., 'slide9.xml'
    if (specificSlide) {
        console.log(`\nDetailed Text Dump for ${specificSlide}:`);
        console.log('--- REFERENCE (A) ---');
        dumpSlideText(dir1, specificSlide);
        console.log('\n--- GENERATED (B) ---');
        dumpSlideText(dir2, specificSlide);
    }

} catch (err) {
    console.error(err);
}
