const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

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
    const texts = content.match(/<a:t>(.*?)<\/a:t>/g) || [];
    console.log(`--- Text Content for ${slideName} ---`);
    texts.forEach(t => {
        const clean = t.replace(/<\/?a:t>/g, '');
        console.log(`[${clean}]`);
    });
    console.log('-----------------------------------');
}

/**
 * Compares two PPTX files and returns structured stats.
 */
function comparePptx(file1, file2) {
    const outDir = path.resolve('temp_compare');
    const dir1 = path.join(outDir, 'A');
    const dir2 = path.join(outDir, 'B');

    // Ensure files exist
    if (!fs.existsSync(file1)) throw new Error(`File not found: ${file1}`);
    if (!fs.existsSync(file2)) throw new Error(`File not found: ${file2}`);

    unzip(file1, dir1);
    unzip(file2, dir2);

    const stats = {
        dimensions: {
            A: getSlideSizes(dir1),
            B: getSlideSizes(dir2)
        },
        media: {
            A: countFiles(path.join(dir1, 'ppt', 'media')),
            B: countFiles(path.join(dir2, 'ppt', 'media'))
        },
        slides: {}
    };

    const slides1 = getSlideContentFeatures(dir1);
    const slides2 = getSlideContentFeatures(dir2);

    Object.keys(slides1).sort().forEach(key => {
        const s1 = slides1[key];
        const s2 = slides2[key];

        if (!s2) {
            stats.slides[key] = { status: 'missing_in_B' };
            return;
        }

        stats.slides[key] = {
            status: 'present',
            diffSize: s2.size - s1.size,
            diffImages: s2.images - s1.images,
            diffShapes: s2.shapes - s1.shapes,
            diffTexts: s2.textBlocks - s1.textBlocks,
            details: { A: s1, B: s2 }
        };
    });

    // Check for slides only in B
    Object.keys(slides2).forEach(key => {
        if (!slides1[key]) {
            stats.slides[key] = { status: 'missing_in_A', details: { B: slides2[key] } };
        }
    });

    // Calculate Similarity Score (Heuristic)
    let totalScore = 0;
    let slideCount = 0;

    Object.keys(stats.slides).forEach(key => {
        const slide = stats.slides[key];
        slideCount++;

        if (slide.status !== 'present') {
            slide.similarity = 0;
            return;
        }

        const s1 = slide.details.A;
        const s2 = slide.details.B;

        // Weights: Texts(40%), Images(30%), Shapes(30%)
        // Score = 1 - (|diff| / max(A, B, 1))

        const scoreText = 1 - (Math.abs(s1.textBlocks - s2.textBlocks) / Math.max(s1.textBlocks, s2.textBlocks, 1));
        const scoreImg = 1 - (Math.abs(s1.images - s2.images) / Math.max(s1.images, s2.images, 1));
        const scoreShape = 1 - (Math.abs(s1.shapes - s2.shapes) / Math.max(s1.shapes, s2.shapes, 1));

        // Ensure non-negative and weighted
        const weightedScore = (Math.max(0, scoreText) * 0.4) + (Math.max(0, scoreImg) * 0.3) + (Math.max(0, scoreShape) * 0.3);
        slide.similarity = parseFloat((weightedScore * 100).toFixed(1));
        totalScore += slide.similarity;
    });

    stats.similarity = slideCount > 0 ? parseFloat((totalScore / slideCount).toFixed(1)) : 0;

    return stats;
}

// CLI Execution Support
if (require.main === module) {
    const file1 = process.argv[2] ? path.resolve(process.argv[2]) : path.resolve('test_data/test1/test1.pptx');
    const file2 = process.argv[3] ? path.resolve(process.argv[3]) : path.resolve('test_output_combined.pptx');

    try {
        console.log(`Comparing:\n A: ${file1}\n B: ${file2}\n`);
        const result = comparePptx(file1, file2);

        console.log(`Dimensions (A): ${result.dimensions.A}`);
        console.log(`Dimensions (B): ${result.dimensions.B}`);
        console.log(`Overall Similarity: ${result.similarity}%`);

        console.log(`Media Files (A): ${result.media.A}`);
        console.log(`Media Files (B): ${result.media.B} (Diff: ${result.media.B - result.media.A})`);

        console.log('\nSlide Content Analysis:');
        Object.keys(result.slides).sort().forEach(key => {
            const data = result.slides[key];
            if (data.status !== 'present') {
                console.log(`Slide ${key}: ${data.status}`);
            } else {
                const s1 = data.details.A;
                const s2 = data.details.B;
                console.log(`Slide ${key}:`);
                console.log(`  Similarity: ${data.similarity}%`);
                console.log(`  Size:     A=${s1.size}, B=${s2.size} (Diff: ${data.diffSize})`);
                console.log(`  Images:   A=${s1.images}, B=${s2.images}`);
                console.log(`  Shapes:   A=${s1.shapes}, B=${s2.shapes}`);
                console.log(`  Groups:   A=${s1.groups}, B=${s2.groups}`);
                console.log(`  Texts:    A=${s1.textBlocks}, B=${s2.textBlocks}`);
                console.log(`  Backgrnd: A=${s1.hasBackground}, B=${s2.hasBackground}`);
            }
        });

        // Detailed Text Dump logic kept for CLI specific request
        const specificSlide = process.argv[4];
        if (specificSlide) {
            const outDir = path.resolve('temp_compare'); // These are still there after comparePptx runs
            const dir1 = path.join(outDir, 'A');
            const dir2 = path.join(outDir, 'B');
            console.log(`\nDetailed Text Dump for ${specificSlide}:`);
            console.log('--- REFERENCE (A) ---');
            dumpSlideText(dir1, specificSlide);
            console.log('\n--- GENERATED (B) ---');
            dumpSlideText(dir2, specificSlide);
        }

    } catch (err) {
        console.error(err);
    }
}

module.exports = { comparePptx };
