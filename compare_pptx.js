const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function unzip(file, dir) {
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    execSync(`unzip -q "${file}" -d "${dir}"`);
}

function getSlideSizes(dir) {
    const presXml = path.join(dir, 'ppt', 'presentation.xml');
    if (!fs.existsSync(presXml)) return { w: 9144000, h: 6858000 }; // Default 10x7.5 inches
    const content = fs.readFileSync(presXml, 'utf8');
    const cxMatch = content.match(/sldSz cx="(\d+)"/);
    const cyMatch = content.match(/sldSz.*?cy="(\d+)"/);

    return {
        w: cxMatch ? parseInt(cxMatch[1]) : 9144000,
        h: cyMatch ? parseInt(cyMatch[1]) : 6858000
    };
}

// Convert EMU to Inches for easier reading
const emuToInch = (emu) => (emu / 914400).toFixed(2);

// Extract elements (shapes, texts, images) from slide XML
function extractElements(xmlPath) {
    if (!fs.existsSync(xmlPath)) return [];
    const content = fs.readFileSync(xmlPath, 'utf8');

    // Simple regex parser for flat structure analysis
    // Note: This is a simplified parser. For production, use a real XML parser.
    const elements = [];

    // Find all shapes/pictures groups
    // <p:sp> is shape/textbox, <p:pic> is picture
    const shapeMatches = content.match(/<p:(sp|pic)>[\s\S]*?<\/p:\1>/g) || [];

    shapeMatches.forEach(shapeXml => {
        const isPic = shapeXml.startsWith('<p:pic');
        const isText = shapeXml.includes('<a:t>');

        // Extract ID
        const idMatch = shapeXml.match(/ id="(\d+)"/);
        const id = idMatch ? idMatch[1] : 'unknown';

        // Extract Position (Off)
        const offMatch = shapeXml.match(/<a:off x="(\d+)" y="(\d+)"/);
        const x = offMatch ? parseInt(offMatch[1]) : 0;
        const y = offMatch ? parseInt(offMatch[2]) : 0;

        // Extract Size (Ext)
        const extMatch = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        const w = extMatch ? parseInt(extMatch[1]) : 0;
        const h = extMatch ? parseInt(extMatch[2]) : 0;

        // Extract Text content
        let text = '';
        if (isText) {
            const textMatches = shapeXml.match(/<a:t>(.*?)<\/a:t>/g) || [];
            text = textMatches.map(t => t.replace(/<\/?a:t>/g, '')).join('');
        }

        elements.push({
            type: isPic ? 'image' : (isText ? 'text' : 'shape'),
            id,
            x, y, w, h,
            text,
            area: w * h
        });
    });

    return elements;
}

// Calculate intersection over union for two matching elements
function calculateIoU(a, b) {
    const xA = Math.max(a.x, b.x);
    const yA = Math.max(a.y, b.y);
    const xB = Math.min(a.x + a.w, b.x + b.w);
    const yB = Math.min(a.y + a.h, b.y + b.h);

    if (xB < xA || yB < yA) return 0;

    const intersection = (xB - xA) * (yB - yA);
    const union = a.area + b.area - intersection;

    return union === 0 ? 0 : intersection / union;
}

// Calculate distance score (0 to 1, 1 is overlap)
function calculatePositionScore(a, b, slideW, slideH) {
    const centerA = { x: a.x + a.w / 2, y: a.y + a.h / 2 };
    const centerB = { x: b.x + b.w / 2, y: b.y + b.h / 2 };

    const dist = Math.sqrt(Math.pow(centerA.x - centerB.x, 2) + Math.pow(centerA.y - centerB.y, 2));
    const maxDist = Math.sqrt(Math.pow(slideW, 2) + Math.pow(slideH, 2)); // Diagonal

    // Normalized distance score: 1 means perfect match, 0 means max distance
    return Math.max(0, 1 - (dist / (maxDist * 0.5))); // 0.5 factor to penalize distance heavily
}

function calculateSizeScore(a, b) {
    if (a.area === 0 && b.area === 0) return 1;
    if (a.area === 0 || b.area === 0) return 0;

    const diff = Math.abs(a.area - b.area);
    const max = Math.max(a.area, b.area);

    return Math.max(0, 1 - (diff / max));
}

// Greedy matching strategy
function matchElements(listA, listB, slideSize) {
    const matches = [];
    const unmatchedA = [...listA];
    const unmatchedB = [...listB];

    // Match logic:
    // 1. For text: match by content similarity first, then position
    // 2. For image/shape: match by position (IoU or Distance)

    // 1. Text Exact Match
    for (let i = unmatchedA.length - 1; i >= 0; i--) {
        const a = unmatchedA[i];
        if (a.type !== 'text') continue;

        let bestMatchIdx = -1;
        let bestDistScore = -1;

        for (let j = 0; j < unmatchedB.length; j++) {
            const b = unmatchedB[j];
            if (b.type !== 'text') continue;

            // Normalize spaces
            const normA = a.text.replace(/\s+/g, '').trim();
            const normB = b.text.replace(/\s+/g, '').trim();

            if (normA && normA === normB) {
                const distScore = calculatePositionScore(a, b, slideSize.w, slideSize.h);
                if (distScore > bestDistScore) {
                    bestDistScore = distScore;
                    bestMatchIdx = j;
                }
            }
        }

        if (bestMatchIdx !== -1) {
            matches.push({ a, b: unmatchedB[bestMatchIdx], score: bestDistScore });
            unmatchedA.splice(i, 1);
            unmatchedB.splice(bestMatchIdx, 1);
        }
    }

    // 2. Remaining items by Position (Nearest Neighbor)
    for (let i = unmatchedA.length - 1; i >= 0; i--) {
        const a = unmatchedA[i];
        let bestMatchIdx = -1;
        let bestDistScore = 0; // Threshold

        for (let j = 0; j < unmatchedB.length; j++) {
            const b = unmatchedB[j];
            if (a.type !== b.type) continue; // Only match same type

            const distScore = calculatePositionScore(a, b, slideSize.w, slideSize.h);
            if (distScore > bestDistScore) {
                bestDistScore = distScore;
                bestMatchIdx = j;
            }
        }

        if (bestMatchIdx !== -1) {
            matches.push({ a, b: unmatchedB[bestMatchIdx], score: bestDistScore });
            unmatchedA.splice(i, 1);
            unmatchedB.splice(bestMatchIdx, 1);
        }
    }

    return { matches, unmatchedA, unmatchedB };
}

function comparePptx(file1, file2) {
    const outDir = path.resolve('temp_compare');
    const dir1 = path.join(outDir, 'A');
    const dir2 = path.join(outDir, 'B');

    // Ensure files exist
    if (!fs.existsSync(file1)) throw new Error(`File not found: ${file1}`);
    if (!fs.existsSync(file2)) throw new Error(`File not found: ${file2}`);

    unzip(file1, dir1);
    unzip(file2, dir2);

    const slideSize = getSlideSizes(dir1);
    const slidesDir1 = path.join(dir1, 'ppt', 'slides');
    const slidesDir2 = path.join(dir2, 'ppt', 'slides');

    if (!fs.existsSync(slidesDir1)) return { similarity: 0, slides: {} };

    const slideFiles = fs.readdirSync(slidesDir1).filter(f => f.endsWith('.xml'));
    const results = { slides: {}, similarity: 0 };

    let totalScore = 0;

    slideFiles.forEach(f => {
        const xml1 = path.join(slidesDir1, f);
        const xml2 = path.join(slidesDir2, f);

        const elementsA = extractElements(xml1);
        const elementsB = extractElements(xml2);

        const { matches, unmatchedA, unmatchedB } = matchElements(elementsA, elementsB, slideSize);

        // Calculate Scores
        // Count Match Score
        const maxCount = Math.max(elementsA.length, elementsB.length, 1);
        const countScore = matches.length / maxCount;

        // Position & Size Accuracy (Average of matches)
        let posSum = 0;
        let sizeSum = 0;

        matches.forEach(m => {
            posSum += calculatePositionScore(m.a, m.b, slideSize.w, slideSize.h);
            sizeSum += calculateSizeScore(m.a, m.b);
        });

        const posScore = matches.length > 0 ? posSum / matches.length : 0;
        const sizeScore = matches.length > 0 ? sizeSum / matches.length : 0;

        // Final Weighted Score
        // Count: 20%, Position: 40%, Size: 40%
        // If there are no matches but elements exist, score is 0.
        // If both are empty, score is 1.
        let slideScore = 0;
        if (elementsA.length === 0 && elementsB.length === 0) {
            slideScore = 100;
        } else {
            const rawScore = (countScore * 0.2) + (posScore * 0.4) + (sizeScore * 0.4);
            slideScore = parseFloat((rawScore * 100).toFixed(1));
        }

        totalScore += slideScore;

        results.slides[f] = {
            similarity: slideScore,
            metrics: {
                countScore: (countScore * 100).toFixed(1),
                posScore: (posScore * 100).toFixed(1),
                sizeScore: (sizeScore * 100).toFixed(1)
            },
            elements: { A: elementsA.length, B: elementsB.length, matched: matches.length }
        };
    });

    results.similarity = slideFiles.length > 0 ? parseFloat((totalScore / slideFiles.length).toFixed(1)) : 0;
    return results;
}

// CLI Execution Support
if (require.main === module) {
    const file1 = process.argv[2] ? path.resolve(process.argv[2]) : path.resolve('test_data/test1/test1.pptx');
    const file2 = process.argv[3] ? path.resolve(process.argv[3]) : path.resolve('test1_output.pptx');

    try {
        console.log(`Comparing:\n A: ${file1}\n B: ${file2}\n`);
        const result = comparePptx(file1, file2);

        console.log(`Overall Similarity: ${result.similarity}%`);
        console.log('\nSlide Details:');

        Object.keys(result.slides).sort().forEach(key => {
            const s = result.slides[key];
            console.log(`Slide ${key}: ${s.similarity}%`);
            console.log(`  Count Match: ${s.metrics.countScore}% (A:${s.elements.A}, B:${s.elements.B}, Match:${s.elements.matched})`);
            console.log(`  Position:    ${s.metrics.posScore}%`);
            console.log(`  Size:        ${s.metrics.sizeScore}%`);
        });

    } catch (err) {
        console.error(err);
    }
}

module.exports = { comparePptx };
