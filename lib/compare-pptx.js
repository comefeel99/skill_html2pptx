/**
 * PPTX Comparison Module
 * Compares two PPTX files and calculates similarity scores
 */

const fs = require('fs');
const path = require('path');
const {
    unzip,
    getSlideSizes,
    extractElements,
    calculatePositionScore,
    calculateSizeScore
} = require('./pptx-utils');

/**
 * Match elements between two slide element lists
 * @param {Array} listA - Elements from reference slide
 * @param {Array} listB - Elements from generated slide
 * @param {object} slideSize - { width, height } in inches
 * @returns {object} { matches, unmatchedA, unmatchedB }
 */
function matchElements(listA, listB, slideSize) {
    const matches = [];
    const unmatchedA = [...listA];
    const unmatchedB = [...listB];

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
                const distScore = calculatePositionScore(a, b, slideSize.width, slideSize.height);
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
        let bestDistScore = 0;

        for (let j = 0; j < unmatchedB.length; j++) {
            const b = unmatchedB[j];
            if (a.type !== b.type) continue;

            const distScore = calculatePositionScore(a, b, slideSize.width, slideSize.height);
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

/**
 * Compare two PPTX files
 * @param {string} file1 - Path to reference PPTX
 * @param {string} file2 - Path to generated PPTX
 * @returns {object} Comparison result with similarity scores
 */
function comparePptx(file1, file2) {
    const outDir = path.resolve(__dirname, '..', 'temp_compare');
    const dir1 = path.join(outDir, 'A');
    const dir2 = path.join(outDir, 'B');

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
        const maxCount = Math.max(elementsA.length, elementsB.length, 1);
        const countScore = matches.length / maxCount;

        let posSum = 0;
        let sizeSum = 0;

        matches.forEach(m => {
            posSum += calculatePositionScore(m.a, m.b, slideSize.width, slideSize.height);
            sizeSum += calculateSizeScore(m.a, m.b);
        });

        const posScore = matches.length > 0 ? posSum / matches.length : 0;
        const sizeScore = matches.length > 0 ? sizeSum / matches.length : 0;

        // Final Weighted Score: Count 20%, Position 40%, Size 40%
        let slideScore = 0;
        if (elementsA.length === 0 && elementsB.length === 0) {
            slideScore = 100;
        } else {
            const rawScore = (countScore * 0.2) + (posScore / 100 * 0.4) + (sizeScore / 100 * 0.4);
            slideScore = parseFloat((rawScore * 100).toFixed(1));
        }

        totalScore += slideScore;

        results.slides[f] = {
            similarity: slideScore,
            metrics: {
                countScore: (countScore * 100).toFixed(1),
                posScore: posScore.toFixed(1),
                sizeScore: sizeScore.toFixed(1)
            },
            elements: { A: elementsA.length, B: elementsB.length, matched: matches.length }
        };
    });

    results.similarity = slideFiles.length > 0 ? parseFloat((totalScore / slideFiles.length).toFixed(1)) : 0;

    // Cleanup temp directory
    if (fs.existsSync(outDir)) {
        fs.rmSync(outDir, { recursive: true, force: true });
    }

    return results;
}

// CLI Execution Support
if (require.main === module) {
    const file1 = process.argv[2] ? path.resolve(process.argv[2]) : path.resolve(__dirname, '..', 'test_data/test1/test1.pptx');
    const file2 = process.argv[3] ? path.resolve(process.argv[3]) : path.resolve(__dirname, '..', 'test1_output.pptx');

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
