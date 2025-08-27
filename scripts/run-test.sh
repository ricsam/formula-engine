#!/bin/bash

# Simple test runner with logs
# Usage: ./scripts/run-test.sh [test-pattern]

TEST_PATTERN="${1:-""}"

echo "🚀 Starting dev server and running tests..."

# Start dev server in background
bun run dev > logs/dev-server.log 2>&1 &
DEV_PID=$!

# Wait for server to start
sleep 3

# Run tests
if [ -n "$TEST_PATTERN" ]; then
    echo "🎯 Running test: $TEST_PATTERN"
    npx playwright test -g "$TEST_PATTERN" --reporter=line > logs/playwright.log 2>&1
    EXIT_CODE=$?
else
    echo "🎯 Running all tests"
    npx playwright test --reporter=line > logs/playwright.log 2>&1
    EXIT_CODE=$?
fi

# Kill dev server
kill $DEV_PID 2>/dev/null

echo ""
echo "📊 RESULTS:"
echo "=========="

# Show debug output
echo ""
echo "🔍 Debug Output:"
grep "🔍 \[DEBUG\]" logs/playwright.log 2>/dev/null || echo "No debug output"

# Show browser activity
echo ""
echo "🖥️ Browser Activity:"
grep "\[browser\]" logs/dev-server.log 2>/dev/null || echo "No browser activity"

# Show test result
echo ""
echo "🎭 Test Result:"
FAILED_COUNT=$(grep -o "[0-9]\+ failed" logs/playwright.log | grep -o "[0-9]\+" || echo "0")
PASSED_COUNT=$(grep -o "[0-9]\+ passed" logs/playwright.log | grep -o "[0-9]\+" || echo "0")
FLAKY_COUNT=$(grep -o "[0-9]\+ flaky" logs/playwright.log | grep -o "[0-9]\+" || echo "0")

if [ "$FAILED_COUNT" -gt 0 ]; then
    echo "❌ $FAILED_COUNT failed"
fi
if [ "$FLAKY_COUNT" -gt 0 ]; then
    echo "⚠️  $FLAKY_COUNT flaky"
fi
if [ "$PASSED_COUNT" -gt 0 ]; then
    echo "✅ $PASSED_COUNT passed"
fi

# Show the raw result line as well
grep -E "failed|passed" logs/playwright.log | tail -1

echo ""
echo "📁 Full logs: logs/playwright.log and logs/dev-server.log"

exit $EXIT_CODE
