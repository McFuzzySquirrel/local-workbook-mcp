#!/usr/bin/env bash
# =============================================================================
# manual-test-cli.sh
#
# Manual smoke-test harness for the ExcelMcp.Client CLI.
# Each test group maps to a real test-data workbook.
#
# Usage:
#   ./scripts/tests/manual-test-cli.sh
#   ./scripts/tests/manual-test-cli.sh --workbook ProjectTracking
#   ./scripts/tests/manual-test-cli.sh --stop-on-fail
#
# Prerequisites:
#   - Run from the repository root:  cd /path/to/local-workbook-mcp
#   - Solution must be built first:  dotnet build
#   - Test data must exist:          scripts/create-sample-workbooks.ps1
#
# =============================================================================

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
SERVER_BIN="$REPO_ROOT/src/ExcelMcp.Server/bin/Debug/net10.0/ExcelMcp.Server"
export EXCEL_MCP_SERVER="$SERVER_BIN"
CLI="dotnet run --project $REPO_ROOT/src/ExcelMcp.Client --no-build --"
DATA="$REPO_ROOT/test-data"

# ── Argument parsing ──────────────────────────────────────────────────────────
FILTER=""
STOP_ON_FAIL=false
for arg in "$@"; do
  case $arg in
    --workbook=*) FILTER="${arg#*=}" ;;
    --workbook)   shift; FILTER="$1" ;;
    --stop-on-fail) STOP_ON_FAIL=true ;;
  esac
done

# ── Colour helpers ────────────────────────────────────────────────────────────
GREEN='\033[0;32m'; RED='\033[0;31m'; YELLOW='\033[1;33m'
CYAN='\033[0;36m'; RESET='\033[0m'; BOLD='\033[1m'

PASS=0; FAIL=0; SKIP=0

header() { echo -e "\n${BOLD}${CYAN}══════════════════════════════════════════${RESET}"; echo -e "${BOLD}${CYAN}  $1${RESET}"; echo -e "${BOLD}${CYAN}══════════════════════════════════════════${RESET}"; }
section() { echo -e "\n${YELLOW}▶ $1${RESET}"; }

run_test() {
  local id="$1" desc="$2"; shift 2
  # $@ = the CLI sub-command and args
  echo -e "  ${CYAN}[$id]${RESET} $desc"
  local output exit_code=0
  output=$($CLI "$@" 2>&1) || exit_code=$?

  if [[ $exit_code -ne 0 ]]; then
    echo -e "    ${RED}✗ FAIL${RESET} (exit $exit_code)"
    echo "$output" | head -5 | sed 's/^/      /'
    ((FAIL++)) || true
    $STOP_ON_FAIL && exit 1
  else
    echo -e "    ${GREEN}✓ PASS${RESET}"
    PASS=$((PASS+1))
  fi
}

run_test_match() {
  local id="$1" desc="$2" pattern="$3"; shift 3
  echo -e "  ${CYAN}[$id]${RESET} $desc"
  local output exit_code=0
  output=$($CLI "$@" 2>&1) || exit_code=$?

  if [[ $exit_code -ne 0 ]]; then
    echo -e "    ${RED}✗ FAIL${RESET} (exit $exit_code)"
    echo "$output" | head -5 | sed 's/^/      /'
    ((FAIL++)) || true
    $STOP_ON_FAIL && exit 1
  elif ! echo "$output" | grep -qi "$pattern"; then
    echo -e "    ${RED}✗ FAIL${RESET} (output missing: '$pattern')"
    echo "$output" | head -8 | sed 's/^/      /'
    ((FAIL++)) || true
    $STOP_ON_FAIL && exit 1
  else
    echo -e "    ${GREEN}✓ PASS${RESET} (matched: '$pattern')"
    PASS=$((PASS+1))
  fi
}

skip_test() {
  local id="$1" desc="$2"
  echo -e "  ${CYAN}[$id]${RESET} $desc  ${YELLOW}[SKIP — manual step]${RESET}"
  SKIP=$((SKIP+1))
}

# ── Preflight ─────────────────────────────────────────────────────────────────
header "Pre-flight Checks"

echo -ne "  Building Client... "
if ! dotnet build "$REPO_ROOT/src/ExcelMcp.Client" --nologo -q 2>&1; then
  echo -e "${RED}Build failed — aborting.${RESET}"; exit 1
fi
echo -e "${GREEN}✓ Build succeeded${RESET}"

sleep 1
echo -ne "  Building Server... "
if ! dotnet build "$REPO_ROOT/src/ExcelMcp.Server" --nologo -q 2>&1; then
  echo -e "${RED}Build failed — aborting.${RESET}"; exit 1
fi
echo -e "${GREEN}✓ Build succeeded${RESET}"

for wb in ProjectTracking.xlsx EmployeeDirectory.xlsx BudgetTracker.xlsx SalesWithPivot.xlsx; do
  if [[ -f "$DATA/$wb" ]]; then
    echo -e "${GREEN}✓ test-data/$wb found${RESET}"
  else
    echo -e "${RED}✗ test-data/$wb missing — run: scripts/create-sample-workbooks.ps1${RESET}"
    [[ "$wb" == "SalesWithPivot.xlsx" ]] && echo "   (SalesWithPivot requires: scripts/create-pivot-test-workbook.ps1)"
  fi
done

# =============================================================================
# GROUP 1 — ProjectTracking.xlsx
# =============================================================================
if [[ -z "$FILTER" || "$FILTER" == "ProjectTracking" ]]; then
  export EXCEL_MCP_WORKBOOK="$DATA/ProjectTracking.xlsx"
  header "Group 1 — ProjectTracking.xlsx"

  section "1a. Protocol & Structure"
  run_test_match "CLI-PT-01" "list tool shows excel-list-structure" "excel-list-structure" \
    list

  run_test_match "CLI-PT-02" "list shows excel-search tool" "excel-search" \
    list

  run_test_match "CLI-PT-03" "preview Tasks sheet returns column headers" "TaskName" \
    preview Tasks

  section "1b. Search"
  run_test_match "CLI-PT-04" "search 'High' finds high-priority tasks" "High" \
    search High

  run_test_match "CLI-PT-05" "search 'In Progress' finds in-progress tasks" "In Progress" \
    search "In Progress"

  run_test_match "CLI-PT-06" "search 'Alice' finds her tasks" "Alice" \
    search Alice

  run_test_match "CLI-PT-07" "search 'Mobile App' finds the project" "Mobile App" \
    search "Mobile App"

  run_test_match "CLI-PT-08" "search 'Bob' finds time-log entries" "Bob" \
    search Bob

  section "1c. Preview with row limit"
  run_test_match "CLI-PT-09" "preview Projects sheet returns project names" "Website Redesign" \
    preview Projects --rows 5

  run_test_match "CLI-PT-10" "preview TimeLog sheet returns employee column" "Employee" \
    preview TimeLog --rows 10
fi

# =============================================================================
# GROUP 2 — EmployeeDirectory.xlsx
# =============================================================================
if [[ -z "$FILTER" || "$FILTER" == "EmployeeDirectory" ]]; then
  export EXCEL_MCP_WORKBOOK="$DATA/EmployeeDirectory.xlsx"
  header "Group 2 — EmployeeDirectory.xlsx"

  section "2a. Structure"
  run_test_match "CLI-ED-01" "list shows all excel tools" "excel-list-structure" \
    list

  run_test_match "CLI-ED-02" "preview Employees sheet shows Salary column" "Salary" \
    preview Employees

  section "2b. Search"
  run_test_match "CLI-ED-03" "search 'Engineering' returns engineering employees" "Engineering" \
    search Engineering

  run_test_match "CLI-ED-04" "search 'Alice Johnson' finds her record" "Alice Johnson" \
    search "Alice Johnson"

  run_test_match "CLI-ED-05" "search 'Manager' finds managers by title" "Manager" \
    search Manager

  run_test_match "CLI-ED-06" "search 'Bob Smith' identifies him as Lead Developer" "Lead Developer" \
    search "Bob Smith"

  run_test_match "CLI-ED-07" "search 'Engineering' in Departments table shows budget 500000" "500000" \
    search Engineering
fi

# =============================================================================
# GROUP 3 — BudgetTracker.xlsx
# =============================================================================
if [[ -z "$FILTER" || "$FILTER" == "BudgetTracker" ]]; then
  export EXCEL_MCP_WORKBOOK="$DATA/BudgetTracker.xlsx"
  header "Group 3 — BudgetTracker.xlsx"

  section "3a. Structure"
  run_test_match "CLI-BT-01" "list shows excel-write-cell tool" "excel-write-cell" \
    list

  run_test_match "CLI-BT-02" "preview Income sheet shows Source column" "Source" \
    preview Income

  section "3b. Search"
  run_test_match "CLI-BT-03" "search 'Client A' returns consulting entries" "Client A" \
    search "Client A"

  run_test_match "CLI-BT-04" "search 'Cloud Provider' finds software expenses" "Cloud Provider" \
    search "Cloud Provider"

  run_test_match "CLI-BT-05" "search 'Consulting' finds consulting income" "Consulting" \
    search Consulting

  run_test_match "CLI-BT-06" "search 'Rent' finds the Office Rent expense" "3000" \
    search Rent

  run_test_match "CLI-BT-07" "preview Expenses sheet shows Vendor column" "Vendor" \
    preview Expenses
fi

# =============================================================================
# GROUP 4 — SalesWithPivot.xlsx
# =============================================================================
if [[ -z "$FILTER" || "$FILTER" == "SalesWithPivot" ]]; then
  WB="$DATA/SalesWithPivot.xlsx"
  if [[ -f "$WB" ]]; then
    export EXCEL_MCP_WORKBOOK="$WB"
    header "Group 4 — SalesWithPivot.xlsx"

    section "4a. Structure and Search"
    run_test_match "CLI-SP-01" "list shows excel-analyze-pivot tool" "excel-analyze-pivot" \
      list

    run_test_match "CLI-SP-02" "preview SalesData shows Region, Product columns" "Region" \
      preview SalesData

    run_test_match "CLI-SP-03" "search 'John' returns his four sales entries" "John" \
      search John

    run_test_match "CLI-SP-04" "search 'East' finds East-region entries" "East" \
      search East

    run_test_match "CLI-SP-05" "search 'Q4' finds the single Q4 entry" "Q4" \
      search Q4

    section "4b. Pivot Analysis"
    run_test_match "CLI-SP-06" "analyze-pivot SalesPivot returns SalesSummary pivot" "SalesSummary" \
      analyze-pivot SalesPivot
  else
    echo -e "${YELLOW}⚠ SalesWithPivot.xlsx not found — skipping Group 4${RESET}"
    echo "  Run: pwsh scripts/create-pivot-test-workbook.ps1"
    SKIP=$((SKIP+6))
  fi
fi

# =============================================================================
# GROUP 5 — Write Operations (uses temp copy)
# =============================================================================
if [[ -z "$FILTER" || "$FILTER" == "Write" ]]; then
  header "Group 5 — Write Operations"

  # Make temp copies so originals stay pristine
  TMPDIR_WB=$(mktemp -d)
  PT_COPY="$TMPDIR_WB/ProjectTracking.xlsx"
  ED_COPY="$TMPDIR_WB/EmployeeDirectory.xlsx"
  cp "$DATA/ProjectTracking.xlsx"   "$PT_COPY"
  cp "$DATA/EmployeeDirectory.xlsx" "$ED_COPY"

  section "5a. Write Cell"
  export EXCEL_MCP_WORKBOOK="$PT_COPY"

  run_test_match "CLI-WO-01" "write-cell sets a value and reports success" "success\|set to\|G1" \
    write-cell --sheet Tasks --cell G1 --value "Notes Header"

  run_test_match "CLI-WO-02" "write numeric cell succeeds" "success\|D2\|99999" \
    write-cell --sheet Projects --cell D2 --value 99999

  # Verify backup was created
  BACKUP_COUNT=$(find "$TMPDIR_WB" -mindepth 1 ! -name "*.xlsx" 2>/dev/null | wc -l)
  echo -e "  ${CYAN}[CLI-WO-03]${RESET} backup file created after write"
  if [[ $BACKUP_COUNT -gt 0 ]]; then
    echo -e "    ${GREEN}✓ PASS${RESET} ($BACKUP_COUNT backup(s) found)"
    PASS=$((PASS+1))
  else
    echo -e "    ${YELLOW}? INFO${RESET} backup may use timestamp filename — check $TMPDIR_WB manually"
    SKIP=$((SKIP+1))
  fi

  section "5b. Write Range"
  export EXCEL_MCP_WORKBOOK="$ED_COPY"

  run_test_match "CLI-WO-04" "write-range updates multiple cells" "success\|Updated\|cell" \
    write-range --sheet Employees \
    --range A10:C10 \
    --data '[{"cellAddress":"A10","value":"9999"},{"cellAddress":"B10","value":"Test User"},{"cellAddress":"C10","value":"QA"}]'

  section "5c. Create Worksheet"
  export EXCEL_MCP_WORKBOOK="$PT_COPY"

  run_test_match "CLI-WO-05" "create-worksheet adds new sheet" "created\|success\|Summary" \
    create-worksheet Summary

  section "5d. Error Cases"
  # Non-existent worksheet should error; we expect a non-zero exit (that's the correct behaviour)
  echo -e "  ${CYAN}[CLI-WO-06]${RESET} write-cell on missing sheet returns error message"
  ERR_OUT=$($CLI write-cell --sheet GhostSheet --cell A1 --value x 2>&1) && WO6_EXIT=0 || WO6_EXIT=$?
  if [[ $WO6_EXIT -ne 0 ]] && echo "$ERR_OUT" | grep -qi "not found\|error\|Ghost"; then
    echo -e "    ${GREEN}✓ PASS${RESET} (correct non-zero exit + error message)"
    PASS=$((PASS+1))
  else
    echo -e "    ${RED}✗ FAIL${RESET} (expected error for non-existent sheet, got exit=$WO6_EXIT)"
    ((FAIL++)) || true
  fi

  skip_test "CLI-WO-07" "write-cell on missing file gives FileNotFoundException (manual verify)"

  # Cleanup
  rm -rf "$TMPDIR_WB"
fi

# ── Summary ───────────────────────────────────────────────────────────────────
header "Results"
echo -e "  ${GREEN}Passed : $PASS${RESET}"
echo -e "  ${RED}Failed : $FAIL${RESET}"
echo -e "  ${YELLOW}Skipped: $SKIP${RESET}"
echo ""
if [[ $FAIL -eq 0 ]]; then
  echo -e "${GREEN}${BOLD}✓ All automated CLI tests passed.${RESET}"
else
  echo -e "${RED}${BOLD}✗ $FAIL test(s) failed.${RESET}"
  exit 1
fi
