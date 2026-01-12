/**
 * Regression Test Runner
 * Executes all test cases and reports results
 */

const fs = require('fs');
const path = require('path');
const { ASSERTION_TYPES, validateTestCase } = require('./test_cases/test_case_schema');
const { comparePptx } = require('./lib/compare-pptx');
const {
    unzip,
    extractTextsFromSlide,
    extractElementCount
} = require('./lib/pptx-utils');

const TEST_CASES_DIR = path.resolve(__dirname, 'test_cases');
const TEST_DATA_DIR = path.resolve(__dirname, 'test_data');
const RESULT_DIR = path.resolve(__dirname, 'test_result');


// 단일 테스트 케이스 실행
function runAssertion(testCase, outputPptxPath, comparisonResult) {
    const { assertion, slideNum } = testCase;
    const tempDir = path.resolve(__dirname, 'temp_regression');

    let result = { passed: false, message: '' };

    try {
        switch (assertion.type) {
            case ASSERTION_TYPES.TEXT_COUNT: {
                unzip(outputPptxPath, tempDir);
                const xmlPath = path.join(tempDir, `ppt/slides/slide${slideNum}.xml`);
                const texts = extractTextsFromSlide(xmlPath);
                const fullText = texts.join('');
                const count = (fullText.match(new RegExp(assertion.content, 'g')) || []).length;

                result.passed = count === assertion.expectedCount;
                result.message = `Found "${assertion.content}" ${count} times (expected: ${assertion.expectedCount})`;
                break;
            }

            case ASSERTION_TYPES.TEXT_EXISTS: {
                unzip(outputPptxPath, tempDir);
                const xmlPath = path.join(tempDir, `ppt/slides/slide${slideNum}.xml`);
                const texts = extractTextsFromSlide(xmlPath);
                const fullText = texts.join('');
                const exists = fullText.includes(assertion.content);

                result.passed = exists;
                result.message = exists
                    ? `Text "${assertion.content}" found`
                    : `Text "${assertion.content}" NOT found`;
                break;
            }

            case ASSERTION_TYPES.TEXT_NOT_EXISTS: {
                unzip(outputPptxPath, tempDir);
                const xmlPath = path.join(tempDir, `ppt/slides/slide${slideNum}.xml`);
                const texts = extractTextsFromSlide(xmlPath);
                const fullText = texts.join('');
                const exists = fullText.includes(assertion.content);

                result.passed = !exists;
                result.message = exists
                    ? `Text "${assertion.content}" still exists (should not)`
                    : `Text "${assertion.content}" correctly absent`;
                break;
            }

            case ASSERTION_TYPES.SIMILARITY_THRESHOLD: {
                if (!comparisonResult || !comparisonResult.slides) {
                    result.message = 'Comparison result not available';
                    break;
                }
                const slideKey = `slide${slideNum}`;
                const slideData = comparisonResult.slides[slideKey];
                if (!slideData) {
                    result.message = `Slide ${slideNum} not found in comparison`;
                    break;
                }
                const similarity = slideData.similarity || 0;
                result.passed = similarity >= assertion.minSimilarity;
                result.message = `Similarity: ${similarity}% (min: ${assertion.minSimilarity}%)`;
                break;
            }

            case ASSERTION_TYPES.ELEMENT_COUNT: {
                unzip(outputPptxPath, tempDir);
                const xmlPath = path.join(tempDir, `ppt/slides/slide${slideNum}.xml`);
                const counts = extractElementCount(xmlPath);
                const tolerance = assertion.tolerance || 0;
                const diff = Math.abs(counts.total - assertion.expectedCount);

                result.passed = diff <= tolerance;
                result.message = `Elements: ${counts.total} (expected: ${assertion.expectedCount}, tolerance: ${tolerance})`;
                break;
            }

            case ASSERTION_TYPES.IMAGE_COUNT: {
                unzip(outputPptxPath, tempDir);
                const xmlPath = path.join(tempDir, `ppt/slides/slide${slideNum}.xml`);
                const counts = extractElementCount(xmlPath);

                result.passed = counts.images === assertion.expectedCount;
                result.message = `Images: ${counts.images} (expected: ${assertion.expectedCount})`;
                break;
            }

            default:
                result.message = `Unknown assertion type: ${assertion.type}`;
        }
    } catch (err) {
        result.message = `Error: ${err.message}`;
    }

    // Cleanup
    if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
    }

    return result;
}

// 모든 테스트 케이스 로드
function loadTestCases() {
    const cases = [];
    const files = fs.readdirSync(TEST_CASES_DIR).filter(f => f.endsWith('.json'));

    for (const file of files) {
        try {
            const content = fs.readFileSync(path.join(TEST_CASES_DIR, file), 'utf8');
            const testCase = JSON.parse(content);
            const validation = validateTestCase(testCase);

            if (validation.valid) {
                cases.push({ ...testCase, _file: file });
            } else {
                console.warn(`Skipping invalid test case ${file}: ${validation.errors.join(', ')}`);
            }
        } catch (err) {
            console.warn(`Failed to load ${file}: ${err.message}`);
        }
    }

    return cases;
}

// 메인 실행
async function runRegressionTests() {
    console.log('\n=== Regression Test Runner ===\n');

    const testCases = loadTestCases();
    console.log(`Loaded ${testCases.length} test case(s)\n`);

    if (testCases.length === 0) {
        console.log('No test cases found.');
        return;
    }

    // 테스트별 출력 파일 및 비교 결과 캐시
    const outputCache = {};
    const comparisonCache = {};

    const results = [];

    for (const testCase of testCases) {
        if (testCase.status === 'skipped') {
            results.push({
                id: testCase.id,
                status: 'SKIPPED',
                message: 'Test case marked as skipped'
            });
            continue;
        }

        const { testDir } = testCase;
        const outputPath = path.resolve(__dirname, `${testDir}_output.pptx`);
        const referencePath = path.join(TEST_DATA_DIR, testDir, `${testDir}.pptx`);

        // 출력 파일 존재 확인
        if (!fs.existsSync(outputPath)) {
            results.push({
                id: testCase.id,
                status: 'ERROR',
                message: `Output file not found: ${outputPath}`
            });
            continue;
        }

        // 비교 결과 캐시
        if (!comparisonCache[testDir] && fs.existsSync(referencePath)) {
            try {
                comparisonCache[testDir] = comparePptx(referencePath, outputPath);
            } catch (err) {
                comparisonCache[testDir] = null;
            }
        }

        // Assertion 실행
        const assertionResult = runAssertion(testCase, outputPath, comparisonCache[testDir]);

        results.push({
            id: testCase.id,
            description: testCase.description,
            testDir: testCase.testDir,
            slideNum: testCase.slideNum,
            status: assertionResult.passed ? 'PASS' : 'FAIL',
            message: assertionResult.message
        });
    }

    // 결과 출력
    console.log('--- Results ---\n');

    const passed = results.filter(r => r.status === 'PASS').length;
    const failed = results.filter(r => r.status === 'FAIL').length;
    const skipped = results.filter(r => r.status === 'SKIPPED').length;
    const errors = results.filter(r => r.status === 'ERROR').length;

    for (const r of results) {
        const icon = r.status === 'PASS' ? '✅' :
            r.status === 'FAIL' ? '❌' :
                r.status === 'SKIPPED' ? '⏭️' : '⚠️';
        console.log(`${icon} [${r.id}] ${r.status}: ${r.message}`);
    }

    console.log(`\n--- Summary ---`);
    console.log(`Total: ${results.length} | ✅ Pass: ${passed} | ❌ Fail: ${failed} | ⏭️ Skipped: ${skipped} | ⚠️ Error: ${errors}`);

    // 결과 저장
    if (!fs.existsSync(RESULT_DIR)) fs.mkdirSync(RESULT_DIR, { recursive: true });

    const resultPath = path.join(RESULT_DIR, 'regression_results.json');
    const resultData = {
        timestamp: new Date().toISOString(),
        summary: { total: results.length, passed, failed, skipped, errors },
        results
    };
    fs.writeFileSync(resultPath, JSON.stringify(resultData, null, 2));
    console.log(`\nResults saved to ${resultPath}`);

    // Exit code for CI
    if (failed > 0 || errors > 0) {
        process.exitCode = 1;
    }
}

runRegressionTests().catch(console.error);
