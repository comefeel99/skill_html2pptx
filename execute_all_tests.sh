#!/bin/bash
echo "Starting all tests..."

echo "--- Running Test 1 ---"
node run_test.js test_data/test1

echo "--- Running Test 2 ---"
node run_test.js test_data/test2

echo "--- Running Test 3 ---"
node run_test.js test_data/test3

echo "--- Running Test 4 ---"
node run_test.js test_data/test4

echo ""
echo "=== Running Regression Tests ==="
node run_regression.js

echo "All tests completed."
