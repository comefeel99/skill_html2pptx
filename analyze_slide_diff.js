const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function unzip(file, dir) {
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    execSync(`unzip -q "${file}" -d "${dir}"`);
}

function extractElements(xmlPath) {
    if (!fs.existsSync(xmlPath)) return [];
    const content = fs.readFileSync(xmlPath, 'utf8');

    // Simple regex parser
    const elements = [];
    // Matches both standard Shapes and Pictures
    const matches = content.match(/<p:(sp|pic)>[\s\S]*?<\/p:\1>/g) || [];

    matches.forEach(shapeXml => {
        const isPic = shapeXml.startsWith('<p:pic');

        // ID
        const idMatch = shapeXml.match(/ id="(\d+)"/);
        const id = idMatch ? idMatch[1] : 'unknown';

        // Position (Off) - EMU
        const offMatch = shapeXml.match(/<a:off x="(\d+)" y="(\d+)"/);
        const x = offMatch ? parseInt(offMatch[1]) : 0;
        const y = offMatch ? parseInt(offMatch[2]) : 0;

        // Size (Ext) - EMU
        const extMatch = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        const w = extMatch ? parseInt(extMatch[1]) : 0;
        const h = extMatch ? parseInt(extMatch[2]) : 0;

        // Text Content
        let text = '';
        if (!isPic) {
            const textMatches = shapeXml.match(/<a:t>(.*?)<\/a:t>/g) || [];
            text = textMatches.map(t => t.replace(/<\/?a:t>/g, '')).join('');
        }

        // Image Rel ID (if pic)
        let imgRelId = '';
        if (isPic) {
            const blipMatch = shapeXml.match(/<a:blip r:embed="(\w+)"/);
            imgRelId = blipMatch ? blipMatch[1] : '';
        }

        elements.push({
            type: isPic ? 'image' : (text ? 'text' : 'shape'),
            id,
            x: x / 914400, // Convert to inches for display
            y: y / 914400,
            w: w / 914400,
            h: h / 914400,
            text,
            rawXml: shapeXml
        });
    });

    return elements;
}

function compare(fileA, fileB) {
    const dirA = 'temp_diff/A';
    const dirB = 'temp_diff/B';

    unzip(fileA, dirA);
    unzip(fileB, dirB);

    const slide1A = path.join(dirA, 'ppt/slides/slide1.xml');
    const slide1B = path.join(dirB, 'ppt/slides/slide1.xml');

    const elemsA = extractElements(slide1A);
    const elemsB = extractElements(slide1B);

    console.log(`\n=== Comparing Slide 1 ===`);
    console.log(`Reference (A): ${elemsA.length} elements`);
    console.log(`Generated (B): ${elemsB.length} elements\n`);

    // Greedy Match based on Text Similarity or Position
    const matched = [];
    const unmatchedA = [...elemsA];
    const unmatchedB = [...elemsB];

    // 1. Match by Text Content (Reference perspective)
    for (let i = unmatchedA.length - 1; i >= 0; i--) {
        const a = unmatchedA[i];
        if (a.type !== 'text') continue;

        let bestIdx = -1;
        let bestScore = -1; // 0 to 1 (1 = exact match)

        for (let j = 0; j < unmatchedB.length; j++) {
            const b = unmatchedB[j];
            if (b.type !== 'text') continue;

            // Normalize text: remove spaces
            const txtA = a.text.replace(/\s+/g, '');
            const txtB = b.text.replace(/\s+/g, '');

            if (txtA === txtB) {
                // Exact match found! Check position closeness
                const dist = Math.sqrt(Math.pow(a.x - b.x, 2) + Math.pow(a.y - b.y, 2));
                const score = 100 - dist; // Prefer closer ones
                if (score > bestScore) {
                    bestScore = score;
                    bestIdx = j;
                }
            } else if (txtA.includes(txtB) || txtB.includes(txtA)) {
                // Partial match
                const score = 50;
                if (score > bestScore) {
                    bestScore = score;
                    bestIdx = j;
                }
            }
        }

        if (bestIdx !== -1) {
            matched.push({ a, b: unmatchedB[bestIdx], reason: 'Text Match' });
            unmatchedA.splice(i, 1);
            unmatchedB.splice(bestIdx, 1);
        }
    }

    // 2. Match by Position (for remaining items, mostly shapes/images)
    for (let i = unmatchedA.length - 1; i >= 0; i--) {
        const a = unmatchedA[i];

        let bestIdx = -1;
        let minDist = 9999;

        for (let j = 0; j < unmatchedB.length; j++) {
            const b = unmatchedB[j];
            if (a.type !== b.type) continue;

            const dist = Math.sqrt(Math.pow(a.x - b.x, 2) + Math.pow(a.y - b.y, 2));
            if (dist < 1.0) { // Only match if reasonably close (within 1 inch)
                if (dist < minDist) {
                    minDist = dist;
                    bestIdx = j;
                }
            }
        }

        if (bestIdx !== -1) {
            matched.push({ a, b: unmatchedB[bestIdx], reason: 'Position Match' });
            unmatchedA.splice(i, 1);
            unmatchedB.splice(bestIdx, 1);
        }
    }

    // Report Differences
    console.log('--- Matched Elements Differences ---');
    matched.sort((m1, m2) => m1.a.y - m2.a.y).forEach(m => {
        const { a, b } = m;
        const xDiff = (b.x - a.x).toFixed(3);
        const yDiff = (b.y - a.y).toFixed(3);
        const wDiff = (b.w - a.w).toFixed(3);
        const hDiff = (b.h - a.h).toFixed(3);

        const isPosDiff = Math.abs(xDiff) > 0.05 || Math.abs(yDiff) > 0.05;
        const isSizeDiff = Math.abs(wDiff) > 0.1 || Math.abs(hDiff) > 0.1;
        const isTextDiff = a.text.replace(/\s+/g, '') !== b.text.replace(/\s+/g, '');

        if (isPosDiff || isSizeDiff || isTextDiff) {
            console.log(`\n[${a.type.toUpperCase()}] "${a.text.substring(0, 30)}${a.text.length > 30 ? '...' : ''}"`);
            if (isPosDiff) console.log(`  Position: A(${a.x.toFixed(2)}, ${a.y.toFixed(2)}) vs B(${b.x.toFixed(2)}, ${b.y.toFixed(2)}) => Diff(x:${xDiff}, y:${yDiff})`);
            if (isSizeDiff) console.log(`  Size:     A(${a.w.toFixed(2)}, ${a.h.toFixed(2)}) vs B(${b.w.toFixed(2)}, ${b.h.toFixed(2)}) => Diff(w:${wDiff}, h:${hDiff})`);
            if (isTextDiff) console.log(`  Content:  "${a.text}" vs "${b.text}"`);
        }
    });

    if (unmatchedA.length > 0) {
        console.log('\n--- Missing in Generated (B) ---');
        unmatchedA.forEach(a => {
            console.log(`[${a.type}] at (${a.x.toFixed(2)}, ${a.y.toFixed(2)}): "${a.text}"`);
        });
    }

    if (unmatchedB.length > 0) {
        console.log('\n--- Extra in Generated (B) ---');
        unmatchedB.forEach(b => {
            console.log(`[${b.type}] at (${b.x.toFixed(2)}, ${b.y.toFixed(2)}): "${b.text}"`);
        });
    }
}

const file1 = 'test_data/test1/test1.pptx';
const file2 = 'test1_output.pptx';

compare(file1, file2);
