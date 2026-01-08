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
    const elements = [];
    const matches = content.match(/<p:(sp|pic)>[\s\S]*?<\/p:\1>/g) || [];

    matches.forEach(shapeXml => {
        const isPic = shapeXml.startsWith('<p:pic');
        const idMatch = shapeXml.match(/ id="(\d+)"/);
        const id = idMatch ? idMatch[1] : 'unknown';
        const offMatch = shapeXml.match(/<a:off x="(\d+)" y="(\d+)"/);
        const x = offMatch ? parseInt(offMatch[1]) : 0;
        const y = offMatch ? parseInt(offMatch[2]) : 0;
        const extMatch = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        const w = extMatch ? parseInt(extMatch[1]) : 0;
        const h = extMatch ? parseInt(extMatch[2]) : 0;
        let text = '';
        if (!isPic) {
            const textMatches = shapeXml.match(/<a:t>(.*?)<\/a:t>/g) || [];
            text = textMatches.map(t => t.replace(/<\/?a:t>/g, '')).join('');
        }
        elements.push({
            type: isPic ? 'image' : (text ? 'text' : 'shape'),
            id, x: x / 914400, y: y / 914400, w: w / 914400, h: h / 914400, text
        });
    });
    return elements;
}

const slideNum = process.argv[2] || '2';
const dirA = 'temp_diff/A';
const dirB = 'temp_diff/B';
unzip('test_data/test1/test1.pptx', dirA);
unzip('test1_output.pptx', dirB);

const slideA = path.join(dirA, `ppt/slides/slide${slideNum}.xml`);
const slideB = path.join(dirB, `ppt/slides/slide${slideNum}.xml`);
const elemsA = extractElements(slideA);
const elemsB = extractElements(slideB);

console.log(`=== Slide ${slideNum} Comparison ===`);
console.log(`Reference (정답): ${elemsA.length} 요소`);
console.log(`Generated (생성): ${elemsB.length} 요소\n`);

console.log('--- 정답 요소 목록 ---');
elemsA.sort((a, b) => a.y - b.y || a.x - b.x).forEach((e, i) => {
    const txt = e.text ? `"${e.text.substring(0, 40)}${e.text.length > 40 ? '...' : ''}"` : '';
    console.log(`${i + 1}. [${e.type.toUpperCase()}] pos(${e.x.toFixed(2)}, ${e.y.toFixed(2)}) size(${e.w.toFixed(2)}x${e.h.toFixed(2)}) ${txt}`);
});

console.log('\n--- 생성 요소 목록 ---');
elemsB.sort((a, b) => a.y - b.y || a.x - b.x).forEach((e, i) => {
    const txt = e.text ? `"${e.text.substring(0, 40)}${e.text.length > 40 ? '...' : ''}"` : '';
    console.log(`${i + 1}. [${e.type.toUpperCase()}] pos(${e.x.toFixed(2)}, ${e.y.toFixed(2)}) size(${e.w.toFixed(2)}x${e.h.toFixed(2)}) ${txt}`);
});

// Match and compare
console.log('\n--- 차이점 분석 ---');
const matched = new Set();
const unmatchedA = [];

elemsA.forEach(a => {
    // Find best match in B by text or position
    let bestMatch = null;
    let bestScore = -1;

    elemsB.forEach((b, idx) => {
        if (matched.has(idx)) return;

        if (a.type === 'text' && b.type === 'text') {
            const txtA = a.text.replace(/\s+/g, '');
            const txtB = b.text.replace(/\s+/g, '');
            if (txtA === txtB) {
                bestMatch = { b, idx, reason: 'exact' };
                bestScore = 100;
            } else if (txtA.includes(txtB) || txtB.includes(txtA)) {
                if (bestScore < 50) { bestMatch = { b, idx, reason: 'partial' }; bestScore = 50; }
            }
        } else if (a.type === b.type) {
            const dist = Math.sqrt(Math.pow(a.x - b.x, 2) + Math.pow(a.y - b.y, 2));
            if (dist < 0.5 && bestScore < 30) {
                bestMatch = { b, idx, reason: 'position' };
                bestScore = 30;
            }
        }
    });

    if (bestMatch) {
        matched.add(bestMatch.idx);
        const b = bestMatch.b;
        const xDiff = Math.abs(b.x - a.x);
        const yDiff = Math.abs(b.y - a.y);
        const wDiff = Math.abs(b.w - a.w);
        const hDiff = Math.abs(b.h - a.h);

        if (xDiff > 0.05 || yDiff > 0.05 || wDiff > 0.1 || hDiff > 0.1) {
            console.log(`\n[${a.type.toUpperCase()}] "${a.text.substring(0, 30)}..."`);
            console.log(`  정답: pos(${a.x.toFixed(2)}, ${a.y.toFixed(2)}) size(${a.w.toFixed(2)}x${a.h.toFixed(2)})`);
            console.log(`  생성: pos(${b.x.toFixed(2)}, ${b.y.toFixed(2)}) size(${b.w.toFixed(2)}x${b.h.toFixed(2)})`);
            console.log(`  차이: x=${(b.x - a.x).toFixed(2)}, y=${(b.y - a.y).toFixed(2)}, w=${(b.w - a.w).toFixed(2)}, h=${(b.h - a.h).toFixed(2)}`);
        }
    } else {
        unmatchedA.push(a);
    }
});

if (unmatchedA.length > 0) {
    console.log('\n--- 생성에서 누락 ---');
    unmatchedA.forEach(a => {
        console.log(`[${a.type.toUpperCase()}] "${a.text}" at (${a.x.toFixed(2)}, ${a.y.toFixed(2)})`);
    });
}

const extraB = elemsB.filter((_, idx) => !matched.has(idx));
if (extraB.length > 0) {
    console.log('\n--- 생성에서 추가 ---');
    extraB.forEach(b => {
        console.log(`[${b.type.toUpperCase()}] "${b.text}" at (${b.x.toFixed(2)}, ${b.y.toFixed(2)})`);
    });
}
