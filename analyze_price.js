const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const dirA = 'temp_diff/test3_ref';
const dirB = 'temp_diff/test3_out';
execSync(`unzip -oq test_data/test3/test3.pptx -d ${dirA}`);
execSync(`unzip -oq test3_output.pptx -d ${dirB}`);

function findShape(xml, searchText) {
    const matches = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
    for (const shape of matches) {
        if (shape.includes(searchText)) {
            const offMatch = shape.match(/<a:off x="(\d+)" y="(\d+)"/);
            const extMatch = shape.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
            const fillMatch = shape.match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/);
            const colorMatch = fillMatch ? fillMatch[0].match(/<a:srgbClr val="([A-Fa-f0-9]+)"/) : null;

            return {
                x: offMatch ? (parseInt(offMatch[1]) / 914400).toFixed(2) : 'N/A',
                y: offMatch ? (parseInt(offMatch[2]) / 914400).toFixed(2) : 'N/A',
                w: extMatch ? (parseInt(extMatch[1]) / 914400).toFixed(2) : 'N/A',
                h: extMatch ? (parseInt(extMatch[2]) / 914400).toFixed(2) : 'N/A',
                hasFill: !!fillMatch,
                fillColor: colorMatch ? colorMatch[1] : 'none'
            };
        }
    }
    return null;
}

const refXml = fs.readFileSync(path.join(dirA, 'ppt/slides/slide2.xml'), 'utf8');
const outXml = fs.readFileSync(path.join(dirB, 'ppt/slides/slide2.xml'), 'utf8');

const refShape = findShape(refXml, '5,000');
const outShape = findShape(outXml, '5,000');

console.log('=== "5,000 ~ 8,000엔대" 분석 ===\n');

console.log('■ 정답 (reference):');
console.log('  위치: (' + refShape?.x + '", ' + refShape?.y + '")');
console.log('  크기: ' + refShape?.w + '" x ' + refShape?.h + '"');
console.log('  배경 채우기: ' + (refShape?.hasFill ? '있음 (#' + refShape?.fillColor + ')' : '없음'));

console.log('\n■ 생성된 출력:');
console.log('  위치: (' + outShape?.x + '", ' + outShape?.y + '")');
console.log('  크기: ' + outShape?.w + '" x ' + outShape?.h + '"');
console.log('  배경 채우기: ' + (outShape?.hasFill ? '있음 (#' + outShape?.fillColor + ')' : '없음'));

console.log('\n■ 차이:');
console.log('  X 차이: ' + (parseFloat(outShape?.x) - parseFloat(refShape?.x)).toFixed(2) + '"');
console.log('  Y 차이: ' + (parseFloat(outShape?.y) - parseFloat(refShape?.y)).toFixed(2) + '"');
