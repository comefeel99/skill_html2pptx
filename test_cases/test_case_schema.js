/**
 * Test Case Schema and Utilities
 * Defines assertion types and validation for regression test cases
 */

const ASSERTION_TYPES = {
    // 특정 텍스트가 정확히 N번 등장하는지 확인
    TEXT_COUNT: 'text_count',
    // 특정 텍스트가 존재하는지 확인
    TEXT_EXISTS: 'text_exists',
    // 특정 텍스트가 존재하지 않는지 확인
    TEXT_NOT_EXISTS: 'text_not_exists',
    // 슬라이드 내 요소 개수 확인
    ELEMENT_COUNT: 'element_count',
    // 요소 위치가 특정 범위 내에 있는지 확인
    ELEMENT_POSITION: 'element_position',
    // 슬라이드 유사도가 최소값 이상인지 확인
    SIMILARITY_THRESHOLD: 'similarity_threshold',
    // 이미지 개수 확인
    IMAGE_COUNT: 'image_count'
};

/**
 * Test Case Schema:
 * {
 *   id: string,              // 고유 식별자 (예: "issue_001")
 *   description: string,     // 이슈 설명
 *   testDir: string,         // 테스트 디렉토리 (예: "test3")
 *   slideNum: number,        // 슬라이드 번호 (1-indexed)
 *   assertion: {
 *     type: ASSERTION_TYPES,
 *     // type별 추가 필드:
 *     // TEXT_COUNT: { content: string, expectedCount: number }
 *     // TEXT_EXISTS: { content: string }
 *     // TEXT_NOT_EXISTS: { content: string }
 *     // ELEMENT_COUNT: { expectedCount: number, tolerance?: number }
 *     // ELEMENT_POSITION: { content: string, x: {min, max}, y: {min, max} }
 *     // SIMILARITY_THRESHOLD: { minSimilarity: number }
 *     // IMAGE_COUNT: { expectedCount: number }
 *   },
 *   status: "active" | "fixed" | "skipped",
 *   createdAt: string,       // ISO date
 *   fixedAt?: string         // ISO date (when fixed)
 * }
 */

function validateTestCase(testCase) {
    const errors = [];

    if (!testCase.id) errors.push('Missing required field: id');
    if (!testCase.description) errors.push('Missing required field: description');
    if (!testCase.testDir) errors.push('Missing required field: testDir');
    if (!testCase.slideNum || testCase.slideNum < 1) errors.push('Invalid slideNum');
    if (!testCase.assertion) errors.push('Missing required field: assertion');

    if (testCase.assertion) {
        const validTypes = Object.values(ASSERTION_TYPES);
        if (!validTypes.includes(testCase.assertion.type)) {
            errors.push(`Invalid assertion type: ${testCase.assertion.type}`);
        }

        // Type-specific validation
        switch (testCase.assertion.type) {
            case ASSERTION_TYPES.TEXT_COUNT:
                if (!testCase.assertion.content) errors.push('TEXT_COUNT requires content');
                if (testCase.assertion.expectedCount === undefined) errors.push('TEXT_COUNT requires expectedCount');
                break;
            case ASSERTION_TYPES.TEXT_EXISTS:
            case ASSERTION_TYPES.TEXT_NOT_EXISTS:
                if (!testCase.assertion.content) errors.push(`${testCase.assertion.type} requires content`);
                break;
            case ASSERTION_TYPES.SIMILARITY_THRESHOLD:
                if (testCase.assertion.minSimilarity === undefined) errors.push('SIMILARITY_THRESHOLD requires minSimilarity');
                break;
            case ASSERTION_TYPES.ELEMENT_COUNT:
            case ASSERTION_TYPES.IMAGE_COUNT:
                if (testCase.assertion.expectedCount === undefined) errors.push(`${testCase.assertion.type} requires expectedCount`);
                break;
        }
    }

    return {
        valid: errors.length === 0,
        errors
    };
}

function createTestCase(options) {
    const testCase = {
        id: options.id,
        description: options.description,
        testDir: options.testDir,
        slideNum: options.slideNum,
        assertion: options.assertion,
        status: options.status || 'active',
        createdAt: new Date().toISOString()
    };

    const validation = validateTestCase(testCase);
    if (!validation.valid) {
        throw new Error(`Invalid test case: ${validation.errors.join(', ')}`);
    }

    return testCase;
}

module.exports = {
    ASSERTION_TYPES,
    validateTestCase,
    createTestCase
};
