#!/bin/bash

# Test runner script that captures both dev server and Playwright test logs
# Usage: ./scripts/test-with-logs.sh [test-pattern]
# Example: ./scripts/test-with-logs.sh "should show formula in formula bar"

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Create logs directory if it doesn't exist
mkdir -p logs

# Get test pattern from command line argument
TEST_PATTERN="${1:-""}"

echo -e "${BLUE}🚀 Starting dev server and Playwright tests...${NC}"

# Start dev server in background and capture its output
echo -e "${YELLOW}📡 Starting dev server...${NC}"
bun run dev > logs/dev-server.log 2>&1 &
DEV_SERVER_PID=$!

# Function to cleanup processes
cleanup() {
    echo -e "\n${YELLOW}🧹 Cleaning up...${NC}"
    if kill -0 $DEV_SERVER_PID 2>/dev/null; then
        kill $DEV_SERVER_PID
        echo -e "${GREEN}✅ Dev server stopped${NC}"
    fi
}

# Set trap to cleanup on interrupt only (not on normal exit)
trap cleanup SIGINT SIGTERM

# Wait for dev server to start (check if port 3000 is available)
echo -e "${YELLOW}⏳ Waiting for dev server to start...${NC}"
for i in {1..30}; do
    if curl -s http://localhost:3000 > /dev/null 2>&1; then
        echo -e "${GREEN}✅ Dev server is running${NC}"
        break
    fi
    if [ $i -eq 30 ]; then
        echo -e "${RED}❌ Dev server failed to start within 30 seconds${NC}"
        exit 1
    fi
    sleep 1
done

# Run Playwright tests and capture output
echo -e "${BLUE}🎭 Running Playwright tests...${NC}"
if [ -n "$TEST_PATTERN" ]; then
    echo -e "${YELLOW}🎯 Running test pattern: $TEST_PATTERN${NC}"
    npx playwright test -g "$TEST_PATTERN" --reporter=line > logs/playwright.log 2>&1
    PLAYWRIGHT_EXIT_CODE=$?
else
    echo -e "${YELLOW}🎯 Running all tests${NC}"
    npx playwright test --reporter=line > logs/playwright.log 2>&1
    PLAYWRIGHT_EXIT_CODE=$?
fi

# Display results
echo -e "\n${BLUE}📊 Test Results:${NC}"
echo -e "${BLUE}=================${NC}"

# Show debug logs first (most important)
echo -e "\n${BLUE}🔍 Debug Output:${NC}"
if grep -q "🔍 \[DEBUG\]" logs/playwright.log; then
    grep "🔍 \[DEBUG\]" logs/playwright.log
else
    echo -e "${YELLOW}No debug output found in test logs${NC}"
fi

# Show relevant dev server logs
echo -e "\n${YELLOW}📡 Dev Server Activity:${NC}"
if grep -q "\[browser\]" logs/dev-server.log; then
    grep "\[browser\]" logs/dev-server.log
else
    echo -e "${YELLOW}No browser activity found${NC}"
fi

# Show test summary
echo -e "\n${YELLOW}🎭 Test Summary:${NC}"
if grep -q "passed\|failed" logs/playwright.log; then
    grep -E "passed|failed" logs/playwright.log | tail -1
else
    echo -e "${YELLOW}No test summary found${NC}"
fi

# Option to show full logs
echo -e "\n${BLUE}💡 Full logs available at:${NC}"
echo -e "  📋 Playwright: logs/playwright.log"
echo -e "  🖥️  Dev Server: logs/dev-server.log"

# Cleanup dev server
cleanup

# Exit with Playwright's exit code
if [ $PLAYWRIGHT_EXIT_CODE -eq 0 ]; then
    echo -e "\n${GREEN}✅ All tests passed!${NC}"
else
    echo -e "\n${RED}❌ Some tests failed (exit code: $PLAYWRIGHT_EXIT_CODE)${NC}"
fi

exit $PLAYWRIGHT_EXIT_CODE
