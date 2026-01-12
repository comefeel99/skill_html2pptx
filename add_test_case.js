#!/usr/bin/env node
/**
 * Add Test Case CLI
 * Easily add new test cases from command line
 * 
 * Usage:
 *   node add_test_case.js --id issue_001 --desc "설명" --test test3 --slide 7 --type text_count --content "텍스트" --count 1
 *   node add_test_case.js --id issue_002 --desc "설명" --test test3 --slide 5 --type text_exists --content "검색할 텍스트"
 *   node add_test_case.js --id issue_003 --desc "설명" --test test1 --slide 1 --type similarity_threshold --min 80
 */

const fs = require('fs');
const path = require('path');
const { createTestCase, ASSERTION_TYPES } = require('./test_cases/test_case_schema');

const TEST_CASES_DIR = path.resolve(__dirname, 'test_cases');

function parseArgs(args) {
    const result = {};
    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        if (arg.startsWith('--')) {
            const key = arg.substring(2);
            const value = args[i + 1] && !args[i + 1].startsWith('--') ? args[++i] : true;
            result[key] = value;
        }
    }
    return result;
}

function printUsage() {
    console.log(`
Add Test Case CLI
=================

Usage:
  node add_test_case.js [options]

Required Options:
  --id <string>       Test case ID (e.g., issue_001)
  --desc <string>     Description
  --test <string>     Test directory (e.g., test1, test3)
  --slide <number>    Slide number (1-indexed)
  --type <string>     Assertion type

Assertion Types & Options:
  text_count          --content <text> --count <number>
  text_exists         --content <text>
  text_not_exists     --content <text>
  element_count       --count <number> [--tolerance <number>]
  image_count         --count <number>
  similarity_threshold --min <number>

Examples:
  # 텍스트가 정확히 1번만 나타나는지 확인
  node add_test_case.js --id issue_001 --desc "가격 중복 방지" \\
    --test test3 --slide 7 --type text_count --content "5,400엔" --count 1

  # 텍스트가 반드시 존재하는지 확인
  node add_test_case.js --id issue_002 --desc "2024 누락 방지" \\
    --test test3 --slide 1 --type text_exists --content "2024"

  # 슬라이드 유사도 최소값 확인
  node add_test_case.js --id issue_003 --desc "슬라이드 품질" \\
    --test test1 --slide 1 --type similarity_threshold --min 75
`);
}

function main() {
    const args = parseArgs(process.argv.slice(2));

    if (args.help || Object.keys(args).length === 0) {
        printUsage();
        return;
    }

    // Validate required args
    const required = ['id', 'desc', 'test', 'slide', 'type'];
    for (const key of required) {
        if (!args[key]) {
            console.error(`Error: Missing required argument: --${key}`);
            printUsage();
            process.exit(1);
        }
    }

    // Build assertion
    let assertion;
    const type = args.type;

    switch (type) {
        case 'text_count':
            if (!args.content || args.count === undefined) {
                console.error('Error: text_count requires --content and --count');
                process.exit(1);
            }
            assertion = {
                type: ASSERTION_TYPES.TEXT_COUNT,
                content: args.content,
                expectedCount: parseInt(args.count)
            };
            break;

        case 'text_exists':
            if (!args.content) {
                console.error('Error: text_exists requires --content');
                process.exit(1);
            }
            assertion = {
                type: ASSERTION_TYPES.TEXT_EXISTS,
                content: args.content
            };
            break;

        case 'text_not_exists':
            if (!args.content) {
                console.error('Error: text_not_exists requires --content');
                process.exit(1);
            }
            assertion = {
                type: ASSERTION_TYPES.TEXT_NOT_EXISTS,
                content: args.content
            };
            break;

        case 'element_count':
            if (args.count === undefined) {
                console.error('Error: element_count requires --count');
                process.exit(1);
            }
            assertion = {
                type: ASSERTION_TYPES.ELEMENT_COUNT,
                expectedCount: parseInt(args.count),
                tolerance: args.tolerance ? parseInt(args.tolerance) : 0
            };
            break;

        case 'image_count':
            if (args.count === undefined) {
                console.error('Error: image_count requires --count');
                process.exit(1);
            }
            assertion = {
                type: ASSERTION_TYPES.IMAGE_COUNT,
                expectedCount: parseInt(args.count)
            };
            break;

        case 'similarity_threshold':
            if (args.min === undefined) {
                console.error('Error: similarity_threshold requires --min');
                process.exit(1);
            }
            assertion = {
                type: ASSERTION_TYPES.SIMILARITY_THRESHOLD,
                minSimilarity: parseInt(args.min)
            };
            break;

        default:
            console.error(`Error: Unknown assertion type: ${type}`);
            console.log('Valid types: text_count, text_exists, text_not_exists, element_count, image_count, similarity_threshold');
            process.exit(1);
    }

    // Create test case
    try {
        const testCase = createTestCase({
            id: args.id,
            description: args.desc,
            testDir: args.test,
            slideNum: parseInt(args.slide),
            assertion
        });

        // Save to file
        const filename = `${args.id}.json`;
        const filepath = path.join(TEST_CASES_DIR, filename);

        if (fs.existsSync(filepath) && !args.force) {
            console.error(`Error: Test case file already exists: ${filename}`);
            console.log('Use --force to overwrite');
            process.exit(1);
        }

        fs.writeFileSync(filepath, JSON.stringify(testCase, null, 2));
        console.log(`✅ Test case created: ${filepath}`);
        console.log(JSON.stringify(testCase, null, 2));

    } catch (err) {
        console.error(`Error: ${err.message}`);
        process.exit(1);
    }
}

main();
