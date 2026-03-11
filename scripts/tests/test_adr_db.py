"""Tests for the EJS ADR Database tool (scripts/adr-db.py)."""

from __future__ import annotations

import json
import sqlite3
import textwrap
import unittest
from pathlib import Path

# Allow importing from the scripts directory
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from importlib import import_module

adr_db = import_module("adr-db")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

SAMPLE_ADR = textwrap.dedent("""\
    ---
    ejs:
      type: journey-adr
      version: 1.1
      adr_id: 0042
      title: Use SQLite for ADR Tracking
      date: 2026-03-02
      status: accepted
      session_id: ejs-session-2026-03-02-01
      session_journey: ejs-docs/journey/2026/ejs-session-2026-03-02-01.md

    actors:
      humans:
        - id: alice
          role: lead-engineer
      agents:
        - id: copilot
          role: coding-agent

    context:
      repo: my-repo
      branch: main
    ---

    # Context

    Agents need a fast way to look up past decisions.

    ---

    # Session Intent

    Evaluate SQLite as an ADR index.

    # Decision

    Adopt SQLite for ADR tracking and agent reference.

    ---

    # Rationale

    SQLite is lightweight, zero-config, and available everywhere.

    ---

    # Consequences

    ### Positive
    - Fast lookup
    - No external dependencies

    ### Negative / Trade-offs
    - Extra sync step needed

    ---

    # Key Learnings

    Agents benefit from structured search over raw file reads.

    ---

    # Agent Guidance

    Prefer database queries over scanning all ADR files.
""")

TEMPLATE_ADR = textwrap.dedent("""\
    ---
    ejs:
      type: journey-adr
      version: 1.1
      adr_id: XXXX
      title: <Short, descriptive title>
      date: YYYY-MM-DD
      status: proposed | accepted | deprecated
    ---

    # Context

    Placeholder template.
""")

SAMPLE_ADR_EXTERNAL = textwrap.dedent("""\
    ---
    ejs:
      type: journey-adr
      version: 1.1
      adr_id: 0099
      title: External ADR in Non-Standard Location
      date: 2026-03-03
      status: proposed
      session_id: ejs-session-2026-03-03-01
      session_journey: ejs-docs/journey/2026/ejs-session-2026-03-03-01.md

    actors:
      humans:
        - id: bob
          role: developer
      agents:
        - id: copilot
          role: coding-agent

    context:
      repo: external-repo
      branch: feature
    ---

    # Context

    This ADR lives outside the canonical ejs-docs/adr directory.

    # Decision

    Place ADR alongside feature code for co-location.

    # Rationale

    Co-location keeps decisions close to the code they affect.

    # Consequences

    ### Positive
    - Better discoverability for feature developers

    ### Negative / Trade-offs
    - Harder to find without recursive scanning

    # Key Learnings

    ADRs can live anywhere in the repo.

    # Agent Guidance

    Always scan the full repo tree for ADR files.
""")

# Plain / Nygard-style ADR — no YAML frontmatter
SAMPLE_PLAIN_ADR = textwrap.dedent("""\
    # ADR-001: Use Markdown for Documentation

    ## Status

    Accepted

    ## Context

    We need a lightweight format for project documentation that is easy to
    read in plain text and renders well on GitHub.

    ## Decision

    We will use Markdown for all project documentation.

    ## Consequences

    Markdown is widely supported and familiar to most developers.
    Some complex layouts may require HTML fallbacks.
""")

SAMPLE_PLAIN_ADR_NUMBERED = textwrap.dedent("""\
    # 2. Use PostgreSQL as Primary Database

    ## Status

    Proposed

    ## Context

    The application needs a reliable relational database.

    ## Decision

    We will use PostgreSQL as the primary database.

    ## Consequences

    Strong ACID compliance and excellent ecosystem support.
""")

SAMPLE_PLAIN_ADR_TEMPLATE = textwrap.dedent("""\
    # ADR-000: ADR Template

    ## Status

    [Proposed | Accepted | Deprecated]

    ## Context

    Describe the context here.

    ## Decision

    Describe the decision here.

    ## Consequences

    Describe the consequences here.
""")


def _write_adr(tmp: Path, filename: str, content: str) -> Path:
    fp = tmp / filename
    fp.write_text(content, encoding="utf-8")
    return fp


class _TempDirMixin:
    """Mixin providing a temporary directory and in-memory database."""

    def setUp(self) -> None:
        import tempfile

        self._tmpdir_obj = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmpdir_obj.name)
        self.adr_dir = self.tmp / "adr"
        self.adr_dir.mkdir()
        self.db_path = self.tmp / "test.db"
        self.conn = adr_db.init_db(self.db_path)

    def tearDown(self) -> None:
        self.conn.close()
        self._tmpdir_obj.cleanup()


# ---------------------------------------------------------------------------
# Test cases
# ---------------------------------------------------------------------------


class TestParseFrontmatter(unittest.TestCase):
    def test_valid_frontmatter(self) -> None:
        fm = adr_db._parse_frontmatter(SAMPLE_ADR)
        # PyYAML (YAML 1.1) interprets 0042 as octal → 34.
        # parse_adr_file works around this by extracting adr_id from raw text.
        self.assertEqual(fm["ejs"]["adr_id"], 34)
        self.assertEqual(fm["ejs"]["status"], "accepted")
        self.assertEqual(fm["actors"]["humans"][0]["id"], "alice")

    def test_no_frontmatter(self) -> None:
        self.assertEqual(adr_db._parse_frontmatter("# Just a heading\n"), {})

    def test_template_skipped(self) -> None:
        fm = adr_db._parse_frontmatter(TEMPLATE_ADR)
        self.assertEqual(fm["ejs"]["adr_id"], "XXXX")


class TestExtractSection(unittest.TestCase):
    def test_decision_section(self) -> None:
        section = adr_db._extract_section(SAMPLE_ADR, "Decision")
        self.assertIn("Adopt SQLite", section)

    def test_context_section(self) -> None:
        section = adr_db._extract_section(SAMPLE_ADR, "Context")
        self.assertIn("fast way to look up", section)

    def test_missing_section(self) -> None:
        self.assertEqual(adr_db._extract_section(SAMPLE_ADR, "Nonexistent"), "")


class TestParseAdrFile(_TempDirMixin, unittest.TestCase):
    def test_valid_file(self) -> None:
        fp = _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        # Temporarily patch _repo_root for relative path calculation
        original = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]
        try:
            record = adr_db.parse_adr_file(fp)
        finally:
            adr_db._repo_root = original
        self.assertIsNotNone(record)
        self.assertEqual(record["adr_id"], "0042")
        self.assertEqual(record["title"], "Use SQLite for ADR Tracking")
        self.assertEqual(record["status"], "accepted")
        self.assertEqual(record["context_repo"], "my-repo")
        self.assertIn("Adopt SQLite", record["decision"])

    def test_template_returns_none(self) -> None:
        fp = _write_adr(self.adr_dir, "0000-template.md", TEMPLATE_ADR)
        original = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]
        try:
            self.assertIsNone(adr_db.parse_adr_file(fp))
        finally:
            adr_db._repo_root = original

    def test_no_frontmatter_returns_none(self) -> None:
        fp = _write_adr(self.adr_dir, "bad.md", "# No frontmatter\n")
        original = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]
        try:
            self.assertIsNone(adr_db.parse_adr_file(fp))
        finally:
            adr_db._repo_root = original


class TestParsePlainAdr(_TempDirMixin, unittest.TestCase):
    """Tests for parsing plain / Nygard-style ADRs without YAML frontmatter."""

    def setUp(self) -> None:
        super().setUp()
        self._orig_root = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]

    def tearDown(self) -> None:
        adr_db._repo_root = self._orig_root
        super().tearDown()

    def test_plain_adr_parsed(self) -> None:
        """A plain Nygard-style ADR is parsed successfully."""
        fp = _write_adr(self.adr_dir, "ADR-001-use-markdown.md", SAMPLE_PLAIN_ADR)
        record = adr_db.parse_adr_file(fp)
        self.assertIsNotNone(record)
        self.assertEqual(record["adr_id"], "ADR-001")
        self.assertEqual(record["title"], "Use Markdown for Documentation")
        self.assertEqual(record["status"], "accepted")
        self.assertIn("lightweight format", record["context_section"])
        self.assertIn("Markdown for all project", record["decision"])
        self.assertIn("widely supported", record["consequences"])

    def test_plain_adr_numbered_heading(self) -> None:
        """A plain ADR with '# 2. Title' heading is parsed."""
        fp = _write_adr(self.adr_dir, "0002-use-postgres.md", SAMPLE_PLAIN_ADR_NUMBERED)
        record = adr_db.parse_adr_file(fp)
        self.assertIsNotNone(record)
        self.assertEqual(record["adr_id"], "ADR-002")
        self.assertEqual(record["title"], "Use PostgreSQL as Primary Database")
        self.assertEqual(record["status"], "proposed")
        self.assertIn("PostgreSQL", record["decision"])

    def test_plain_template_skipped(self) -> None:
        """A plain ADR-000 template is skipped."""
        fp = _write_adr(self.adr_dir, "ADR-000-template.md", SAMPLE_PLAIN_ADR_TEMPLATE)
        record = adr_db.parse_adr_file(fp)
        self.assertIsNone(record)

    def test_plain_adr_no_heading_skipped(self) -> None:
        """A file with no heading is skipped."""
        fp = _write_adr(self.adr_dir, "notes.md", "Some notes without headings.\n")
        record = adr_db.parse_adr_file(fp)
        self.assertIsNone(record)

    def test_plain_adr_no_decision_or_context_skipped(self) -> None:
        """A markdown file with a heading but no Decision/Context sections is skipped."""
        content = "# Some Title\n\n## Introduction\n\nJust some text.\n"
        fp = _write_adr(self.adr_dir, "readme.md", content)
        record = adr_db.parse_adr_file(fp)
        self.assertIsNone(record)

    def test_plain_adr_sync_into_db(self) -> None:
        """Plain ADRs found during sync are inserted into the database.

        Uses an EJS-format ADR (SAMPLE_ADR) alongside a plain ADR to verify
        both formats coexist in the same database after sync.
        """
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        other_dir = self.tmp / "docs" / "adr"
        other_dir.mkdir(parents=True)
        _write_adr(other_dir, "ADR-001-use-markdown.md", SAMPLE_PLAIN_ADR)

        rc = adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT adr_id FROM adrs ORDER BY adr_id").fetchall()
        ids = [r["adr_id"] for r in rows]
        self.assertEqual(len(ids), 2)
        self.assertIn("0042", ids)
        self.assertIn("ADR-001", ids)

    def test_plain_adr_ejs_still_preferred(self) -> None:
        """EJS-format ADRs are still parsed with the EJS parser (not the fallback)."""
        fp = _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        record = adr_db.parse_adr_file(fp)
        self.assertIsNotNone(record)
        # EJS parser fills session_id; plain parser leaves it empty
        self.assertEqual(record["session_id"], "ejs-session-2026-03-02-01")
        self.assertEqual(record["context_repo"], "my-repo")


class TestDatabaseSchema(_TempDirMixin, unittest.TestCase):
    def test_tables_created(self) -> None:
        tables = {
            row[0]
            for row in self.conn.execute(
                "SELECT name FROM sqlite_master WHERE type IN ('table', 'trigger')"
            )
        }
        self.assertIn("adrs", tables)
        self.assertIn("adrs_fts", tables)
        self.assertIn("adrs_ai", tables)
        self.assertIn("adrs_ad", tables)
        self.assertIn("adrs_au", tables)
        self.assertIn("journeys", tables)
        self.assertIn("journeys_fts", tables)
        self.assertIn("journeys_ai", tables)
        self.assertIn("journeys_ad", tables)
        self.assertIn("journeys_au", tables)


class TestSync(_TempDirMixin, unittest.TestCase):
    def setUp(self) -> None:
        super().setUp()
        self._orig_root = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]

    def tearDown(self) -> None:
        adr_db._repo_root = self._orig_root
        super().tearDown()

    def test_sync_inserts_records(self) -> None:
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        rc = adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT * FROM adrs").fetchall()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["adr_id"], "0042")

    def test_sync_skips_template(self) -> None:
        _write_adr(self.adr_dir, "0000-template.md", TEMPLATE_ADR)
        adr_db.cmd_sync(self.conn, self.adr_dir)
        rows = self.conn.execute("SELECT * FROM adrs").fetchall()
        self.assertEqual(len(rows), 0)

    def test_sync_upserts_on_change(self) -> None:
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        adr_db.cmd_sync(self.conn, self.adr_dir)

        updated = SAMPLE_ADR.replace("accepted", "deprecated")
        _write_adr(self.adr_dir, "0042-test.md", updated)
        adr_db.cmd_sync(self.conn, self.adr_dir)

        row = self.conn.execute("SELECT status FROM adrs WHERE adr_id = '0042'").fetchone()
        self.assertEqual(row["status"], "deprecated")

    def test_sync_removes_stale(self) -> None:
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(self.conn.execute("SELECT count(*) FROM adrs").fetchone()[0], 1)

        # Remove the file and re-sync
        (self.adr_dir / "0042-test.md").unlink()
        adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(self.conn.execute("SELECT count(*) FROM adrs").fetchone()[0], 0)

    def test_sync_missing_dir(self) -> None:
        rc = adr_db.cmd_sync(self.conn, self.tmp / "nonexistent")
        self.assertEqual(rc, 1)


class TestAdrDiscovery(_TempDirMixin, unittest.TestCase):
    """Tests for discovering ADR files outside the canonical adr directory."""

    def setUp(self) -> None:
        super().setUp()
        self._orig_root = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]

    def tearDown(self) -> None:
        adr_db._repo_root = self._orig_root
        super().tearDown()

    def test_discovers_adr_outside_adr_dir(self) -> None:
        """ADR placed in a non-standard directory is discovered and synced."""
        # Place one ADR in the canonical dir
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        # Place another ADR in a completely different directory
        other_dir = self.tmp / "docs" / "decisions"
        other_dir.mkdir(parents=True)
        _write_adr(other_dir, "0099-external.md", SAMPLE_ADR_EXTERNAL)

        rc = adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT adr_id FROM adrs ORDER BY adr_id").fetchall()
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["adr_id"], "0042")
        self.assertEqual(rows[1]["adr_id"], "0099")

    def test_discovers_adr_in_nested_subdirectory(self) -> None:
        """ADR deeply nested under the repo root is still found."""
        nested = self.tmp / "src" / "feature" / "adr"
        nested.mkdir(parents=True)
        _write_adr(nested, "0099-nested.md", SAMPLE_ADR_EXTERNAL)

        rc = adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT adr_id FROM adrs").fetchall()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["adr_id"], "0099")

    def test_skips_git_directory(self) -> None:
        """Files inside .git should not be scanned."""
        git_dir = self.tmp / ".git" / "refs"
        git_dir.mkdir(parents=True)
        _write_adr(git_dir, "0099-gitfile.md", SAMPLE_ADR_EXTERNAL)

        rc = adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT adr_id FROM adrs").fetchall()
        self.assertEqual(len(rows), 0)

    def test_skips_node_modules(self) -> None:
        """Files inside node_modules should not be scanned."""
        nm_dir = self.tmp / "node_modules" / "some-pkg"
        nm_dir.mkdir(parents=True)
        _write_adr(nm_dir, "0099-dep.md", SAMPLE_ADR_EXTERNAL)

        rc = adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT adr_id FROM adrs").fetchall()
        self.assertEqual(len(rows), 0)

    def test_no_duplicates_when_adr_in_canonical_dir(self) -> None:
        """An ADR in the canonical dir should not be synced twice."""
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)

        rc = adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT adr_id FROM adrs").fetchall()
        self.assertEqual(len(rows), 1)

    def test_stale_removal_covers_external_adrs(self) -> None:
        """Stale external ADRs are removed when files are deleted."""
        other_dir = self.tmp / "docs"
        other_dir.mkdir(parents=True)
        _write_adr(other_dir, "0099-external.md", SAMPLE_ADR_EXTERNAL)

        adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(self.conn.execute("SELECT count(*) FROM adrs").fetchone()[0], 1)

        # Remove the external file and re-sync
        (other_dir / "0099-external.md").unlink()
        adr_db.cmd_sync(self.conn, self.adr_dir)
        self.assertEqual(self.conn.execute("SELECT count(*) FROM adrs").fetchone()[0], 0)

    def test_sync_output_reports_external_count(self) -> None:
        """Sync output message mentions how many ADRs were found externally."""
        import io
        from contextlib import redirect_stdout

        other_dir = self.tmp / "docs"
        other_dir.mkdir(parents=True)
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        _write_adr(other_dir, "0099-external.md", SAMPLE_ADR_EXTERNAL)

        buf = io.StringIO()
        with redirect_stdout(buf):
            adr_db.cmd_sync(self.conn, self.adr_dir)
        output = buf.getvalue()
        self.assertIn("Synced 2 ADR(s)", output)
        self.assertIn("1 found outside", output)

    def test_discover_adr_files_helper(self) -> None:
        """_discover_adr_files returns files from both canonical and external locations."""
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        other_dir = self.tmp / "other"
        other_dir.mkdir()
        _write_adr(other_dir, "0099-other.md", SAMPLE_ADR_EXTERNAL)

        files = adr_db._discover_adr_files(self.tmp, self.adr_dir)
        # Should contain both files (though non-ADR .md files may also appear)
        paths_str = [str(f) for f in files]
        self.assertTrue(any("0042-test.md" in p for p in paths_str))
        self.assertTrue(any("0099-other.md" in p for p in paths_str))


class TestFullTextSearch(_TempDirMixin, unittest.TestCase):
    def setUp(self) -> None:
        super().setUp()
        self._orig_root = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        adr_db.cmd_sync(self.conn, self.adr_dir)

    def tearDown(self) -> None:
        adr_db._repo_root = self._orig_root
        super().tearDown()

    def test_search_finds_match(self) -> None:
        rows = self.conn.execute(
            "SELECT adr_id FROM adrs_fts WHERE adrs_fts MATCH 'SQLite'"
        ).fetchall()
        self.assertEqual(len(rows), 1)

    def test_search_no_match(self) -> None:
        rows = self.conn.execute(
            "SELECT adr_id FROM adrs_fts WHERE adrs_fts MATCH 'nonexistent_xyzzy'"
        ).fetchall()
        self.assertEqual(len(rows), 0)


class TestCLICommands(_TempDirMixin, unittest.TestCase):
    def setUp(self) -> None:
        super().setUp()
        self._orig_root = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        adr_db.cmd_sync(self.conn, self.adr_dir)

    def tearDown(self) -> None:
        adr_db._repo_root = self._orig_root
        super().tearDown()

    def test_cmd_list(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_list(self.conn)
        self.assertEqual(rc, 0)
        self.assertIn("0042", buf.getvalue())
        self.assertIn("Use SQLite", buf.getvalue())

    def test_cmd_get_found(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_get(self.conn, "0042")
        self.assertEqual(rc, 0)
        self.assertIn("Use SQLite for ADR Tracking", buf.getvalue())

    def test_cmd_get_not_found(self) -> None:
        rc = adr_db.cmd_get(self.conn, "9999")
        self.assertEqual(rc, 1)

    def test_cmd_summary(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_summary(self.conn)
        self.assertEqual(rc, 0)
        output = buf.getvalue()
        self.assertIn("ADR 0042", output)
        self.assertIn("accepted", output)

    def test_cmd_search(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_search(self.conn, "SQLite")
        self.assertEqual(rc, 0)
        self.assertIn("0042", buf.getvalue())


class TestMainEntryPoint(_TempDirMixin, unittest.TestCase):
    def setUp(self) -> None:
        super().setUp()
        self._orig_root = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)

    def tearDown(self) -> None:
        adr_db._repo_root = self._orig_root
        super().tearDown()

    def test_main_sync_and_list(self) -> None:
        import io
        from contextlib import redirect_stdout

        db = str(self.db_path)
        adr = str(self.adr_dir)

        # sync
        rc = adr_db.main(["--db", db, "--adr-dir", adr, "sync"])
        self.assertEqual(rc, 0)

        # list
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.main(["--db", db, "--adr-dir", adr, "list"])
        self.assertEqual(rc, 0)
        self.assertIn("0042", buf.getvalue())

    def test_main_no_command(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.main([])
        self.assertEqual(rc, 1)


# ---------------------------------------------------------------------------
# Journey test fixtures and helpers
# ---------------------------------------------------------------------------

SAMPLE_JOURNEY_FRONTMATTER = textwrap.dedent("""\
    ---
    session_id: ejs-session-2026-02-10-01
    author: github-copilot
    date: 2026-02-10
    repo: Engineering-Journey-System
    branch: copilot/capture-sub-agent-decisions
    agents_involved:
      - GitHub Copilot (main agent)
      - explore agent (codebase analysis)
    decision_detected: true
    adr_links:
      - ../../adr/0012-sub-agent-decision-capture-protocol.md
    tags:
      - ejs
      - multi-agent
    refs:
      - .github/agents/ejs-journey.agent.md
    ---

    # Problem / Intent

    Ensure sub-agents capture their decisions and collaboration.

    # Interaction Summary (Required)

    - Human: Asked how to ensure sub-agent decisions are captured
      - Agent: Explored the repository, identified the gap
      - Outcome: Implemented sub-agent contribution protocol

    # Decisions Made

    - Decision: Add Sub-Agent Contributions section
      - Reason: Sub-agent decisions were being lost
      - Impact: Better traceability

    # Key Learnings

    Sub-agent decisions are a distinct data category.

    # Future Agent Guidance

    Prefer recording sub-agent decisions in structured sections.
""")

SAMPLE_JOURNEY_PLAIN = textwrap.dedent("""\
    session_id: ejs-session-2026-03-02-01
    author: github-copilot
    date: 2026-03-02
    repo: Engineering-Journey-System
    branch: copilot/add-sqlite
    agents_involved: [github-copilot, explore-agent]
    decision_detected: false
    adr_links: []
    tags: [sqlite, adr-tracking]
    refs: []

    # Problem / Intent

    Create a SQLite-backed tool for ADR tracking.

    # Interaction Summary (Required)

    - Human: Requested SQLite implementation
      - Agent: Implemented CLI tool with FTS5
      - Outcome: Tool working with 5 commands

    # Key Learnings

    YAML octal parsing can silently corrupt IDs.

    # Future Agent Guidance

    Prefer database queries over file scanning.
""")


class _JourneyTempDirMixin:
    """Mixin providing a temp dir with both adr and journey subdirs."""

    def setUp(self) -> None:
        import tempfile

        self._tmpdir_obj = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmpdir_obj.name)
        self.adr_dir = self.tmp / "adr"
        self.adr_dir.mkdir()
        self.journey_dir = self.tmp / "journey" / "2026"
        self.journey_dir.mkdir(parents=True)
        # Also add a _templates dir to ensure it's skipped
        (self.tmp / "journey" / "_templates").mkdir()
        self.db_path = self.tmp / "test.db"
        self.conn = adr_db.init_db(self.db_path)
        self._orig_root = adr_db._repo_root
        adr_db._repo_root = lambda: self.tmp  # type: ignore[assignment]

    def tearDown(self) -> None:
        adr_db._repo_root = self._orig_root
        self.conn.close()
        self._tmpdir_obj.cleanup()


def _write_journey(d: Path, filename: str, content: str) -> Path:
    fp = d / filename
    fp.write_text(content, encoding="utf-8")
    return fp


# ---------------------------------------------------------------------------
# Journey tests
# ---------------------------------------------------------------------------


class TestParseJourneyFile(_JourneyTempDirMixin, unittest.TestCase):
    def test_frontmatter_journey(self) -> None:
        fp = _write_journey(self.journey_dir, "ejs-session-2026-02-10-01.md",
                            SAMPLE_JOURNEY_FRONTMATTER)
        record = adr_db.parse_journey_file(fp)
        self.assertIsNotNone(record)
        self.assertEqual(record["session_id"], "ejs-session-2026-02-10-01")
        self.assertEqual(record["date"], "2026-02-10")
        self.assertEqual(record["decision_detected"], "true")
        self.assertIn("sub-agents capture", record["problem_intent"])

    def test_plain_metadata_journey(self) -> None:
        fp = _write_journey(self.journey_dir, "ejs-session-2026-03-02-01.md",
                            SAMPLE_JOURNEY_PLAIN)
        record = adr_db.parse_journey_file(fp)
        self.assertIsNotNone(record)
        self.assertEqual(record["session_id"], "ejs-session-2026-03-02-01")
        self.assertEqual(record["date"], "2026-03-02")
        self.assertIn("SQLite-backed tool", record["problem_intent"])

    def test_no_session_id_returns_none(self) -> None:
        fp = _write_journey(self.journey_dir, "bad.md", "# Just a heading\n")
        self.assertIsNone(adr_db.parse_journey_file(fp))

    def test_template_skipped(self) -> None:
        template = "session_id:\nauthor:\ndate:\n\n# Problem / Intent\nDescribe.\n"
        fp = _write_journey(self.journey_dir, "template.md", template)
        self.assertIsNone(adr_db.parse_journey_file(fp))


class TestJourneySync(_JourneyTempDirMixin, unittest.TestCase):
    def test_sync_inserts_journeys(self) -> None:
        _write_journey(self.journey_dir, "ejs-session-2026-02-10-01.md",
                       SAMPLE_JOURNEY_FRONTMATTER)
        _write_journey(self.journey_dir, "ejs-session-2026-03-02-01.md",
                       SAMPLE_JOURNEY_PLAIN)
        rc = adr_db._sync_journeys(self.conn, self.tmp / "journey")
        self.assertEqual(rc, 0)
        rows = self.conn.execute("SELECT * FROM journeys").fetchall()
        self.assertEqual(len(rows), 2)

    def test_sync_removes_stale_journeys(self) -> None:
        _write_journey(self.journey_dir, "ejs-session-2026-02-10-01.md",
                       SAMPLE_JOURNEY_FRONTMATTER)
        adr_db._sync_journeys(self.conn, self.tmp / "journey")
        self.assertEqual(
            self.conn.execute("SELECT count(*) FROM journeys").fetchone()[0], 1)

        (self.journey_dir / "ejs-session-2026-02-10-01.md").unlink()
        adr_db._sync_journeys(self.conn, self.tmp / "journey")
        self.assertEqual(
            self.conn.execute("SELECT count(*) FROM journeys").fetchone()[0], 0)

    def test_sync_skips_templates_dir(self) -> None:
        template_dir = self.tmp / "journey" / "_templates"
        _write_journey(template_dir, "journey-template.md",
                       "session_id: test\n# Problem / Intent\nTemplate.\n")
        adr_db._sync_journeys(self.conn, self.tmp / "journey")
        rows = self.conn.execute("SELECT * FROM journeys").fetchall()
        self.assertEqual(len(rows), 0)

    def test_cmd_sync_includes_journeys(self) -> None:
        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        _write_journey(self.journey_dir, "ejs-session-2026-03-02-01.md",
                       SAMPLE_JOURNEY_PLAIN)
        rc = adr_db.cmd_sync(self.conn, self.adr_dir, self.tmp / "journey")
        self.assertEqual(rc, 0)
        self.assertEqual(
            self.conn.execute("SELECT count(*) FROM adrs").fetchone()[0], 1)
        self.assertEqual(
            self.conn.execute("SELECT count(*) FROM journeys").fetchone()[0], 1)


class TestJourneyFTS(_JourneyTempDirMixin, unittest.TestCase):
    def setUp(self) -> None:
        super().setUp()
        _write_journey(self.journey_dir, "ejs-session-2026-03-02-01.md",
                       SAMPLE_JOURNEY_PLAIN)
        adr_db._sync_journeys(self.conn, self.tmp / "journey")

    def test_fts_finds_journey(self) -> None:
        rows = self.conn.execute(
            "SELECT session_id FROM journeys_fts WHERE journeys_fts MATCH 'SQLite'"
        ).fetchall()
        self.assertEqual(len(rows), 1)

    def test_fts_no_match(self) -> None:
        rows = self.conn.execute(
            "SELECT session_id FROM journeys_fts WHERE journeys_fts MATCH 'nonexistent_xyzzy'"
        ).fetchall()
        self.assertEqual(len(rows), 0)


class TestJourneyCLICommands(_JourneyTempDirMixin, unittest.TestCase):
    def setUp(self) -> None:
        super().setUp()
        _write_journey(self.journey_dir, "ejs-session-2026-03-02-01.md",
                       SAMPLE_JOURNEY_PLAIN)
        adr_db._sync_journeys(self.conn, self.tmp / "journey")

    def test_cmd_list_journeys(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_list_journeys(self.conn)
        self.assertEqual(rc, 0)
        self.assertIn("ejs-session-2026-03-02-01", buf.getvalue())

    def test_cmd_get_journey_found(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_get_journey(self.conn, "ejs-session-2026-03-02-01")
        self.assertEqual(rc, 0)
        self.assertIn("SQLite-backed tool", buf.getvalue())

    def test_cmd_get_journey_not_found(self) -> None:
        rc = adr_db.cmd_get_journey(self.conn, "nonexistent")
        self.assertEqual(rc, 1)

    def test_cmd_summary_journeys(self) -> None:
        import io
        from contextlib import redirect_stdout

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_summary_journeys(self.conn)
        self.assertEqual(rc, 0)
        output = buf.getvalue()
        self.assertIn("ejs-session-2026-03-02-01", output)
        self.assertIn("SQLite-backed tool", output)

    def test_search_across_both(self) -> None:
        import io
        from contextlib import redirect_stdout

        _write_adr(self.adr_dir, "0042-test.md", SAMPLE_ADR)
        adr_db._sync_adrs(self.conn, self.adr_dir)

        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = adr_db.cmd_search(self.conn, "SQLite")
        self.assertEqual(rc, 0)
        output = buf.getvalue()
        self.assertIn("[ADR 0042]", output)
        self.assertIn("[Journey ejs-session-2026-03-02-01]", output)


if __name__ == "__main__":
    unittest.main()
