const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

function unzip(file, dir) {
    if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    execSync(`unzip -q "${file}" -d "${dir}"`);
}

const dirB = 'temp_diff/test3_out';
unzip('test3_output.pptx', dirB);
const xml = fs.readFileSync(path.join(dirB, 'ppt/slides/slide1.xml'), 'utf8');

const matches = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
const elements = matches.map((shape, idx) => {
    const offMatch = shape.match(/<a:off x="(\d+)" y="(\d+)"/);
    const extMatch = shape.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
    const textMatches = shape.match(/<a:t>(.*?)<\/a:t>/g) || [];
    const text = textMatches.map(t => t.replace(/<\/?a:t>/g, '')).join('');

    const x = offMatch ? parseInt(offMatch[1]) / 914400 : 0;
    const y = offMatch ? parseInt(offMatch[2]) / 914400 : 0;
    const w = extMatch ? parseInt(extMatch[1]) / 914400 : 0;
    const h = extMatch ? parseInt(extMatch[2]) / 914400 : 0;

    return { idx, x, y, w, h, right: x + w, bottom: y + h, text: text.substring(0, 40) };
}).filter(e => e.text);

console.log('All text elements found:', elements.map(e => e.text));
const target = elements.find(e => e.text.includes('2024'));
console.log('=== "이세탄 신주쿠점 한정 케이크" 주변 공간 분석 ===\n');
console.log('대상 요소:');
console.log(`  위치: (${target.x.toFixed(2)}, ${target.y.toFixed(2)}) 크기: ${target.w.toFixed(2)} x ${target.h.toFixed(2)}`);
console.log(`  오른쪽 끝: ${target.right.toFixed(2)}"`);

// 같은 Y 범위에 있는 요소들
const sameRow = elements.filter(e => {
    if (e.idx === target.idx) return false;
    const yOverlap = !(e.bottom < target.y || e.y > target.bottom);
    return yOverlap;
});

console.log(`\n같은 행의 다른 요소들:`);
sameRow.sort((a, b) => a.x - b.x).forEach(e => {
    const gap = e.x - target.right;
    console.log(`  X:${e.x.toFixed(2)}" 간격:${gap.toFixed(2)}" 텍스트: "${e.text}"`);
});

const rightElements = sameRow.filter(e => e.x > target.x);
const nearestRight = rightElements.length > 0 ? Math.min(...rightElements.map(e => e.x)) : 13.33;
const availableSpace = nearestRight - target.right;

console.log(`\n✅ 확장 가능 공간 분석:`);
console.log(`   현재 오른쪽 끝: ${target.right.toFixed(2)}"`);
console.log(`   가장 가까운 오른쪽 요소: ${nearestRight.toFixed(2)}"`);
console.log(`   사용 가능한 공간: ${availableSpace.toFixed(2)}"`);
console.log(`   필요한 버퍼 (20%): ${(target.w * 0.20).toFixed(2)}"`);
console.log(`   판정: ${availableSpace >= target.w * 0.20 ? '✓ 확장 가능' : '✗ 겹침 주의'}`);
