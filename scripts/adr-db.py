#!/usr/bin/env python3
"""
EJS Database — SQLite-backed index for ADRs and Session Journeys.

Parses ADR and journey markdown files from ejs-docs/, extracts metadata and
key content sections, and stores them in a local SQLite database for fast
agent-friendly querying when full-file context would be too expensive.

Usage:
    python scripts/adr-db.py sync                   # Parse ADR + journey files → database
    python scripts/adr-db.py list                   # List all ADRs (compact)
    python scripts/adr-db.py get <adr_id>           # Full details for one ADR
    python scripts/adr-db.py search <query>         # Full-text search across ADRs and journeys
    python scripts/adr-db.py summary                # Agent-friendly compact summary of ADRs
    python scripts/adr-db.py list-journeys          # List all journeys (compact)
    python scripts/adr-db.py get-journey <id>       # Full details for one journey
    python scripts/adr-db.py summary-journeys       # Agent-friendly compact summary of journeys

The database is stored at <repo_root>/.ejs.db (gitignored).
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore[assignment]


# ---------------------------------------------------------------------------
# Path helpers
# ---------------------------------------------------------------------------

def _repo_root() -> Path:
    """Return the repository root (parent of scripts/)."""
    return Path(__file__).resolve().parent.parent


def _default_db_path() -> Path:
    return _repo_root() / ".ejs.db"


def _default_adr_dir() -> Path:
    return _repo_root() / "ejs-docs" / "adr"


def _default_journey_dir() -> Path:
    return _repo_root() / "ejs-docs" / "journey"


# Directories to skip when recursively scanning the repo for ADR files.
_SCAN_SKIP_DIRS = frozenset({
    ".git", "node_modules", "dist", ".parcel-cache", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".tox", ".venv", "venv",
    "_templates",
})


def _discover_adr_files(repo_root: Path, adr_dir: Path) -> list[Path]:
    """Discover all ADR markdown files across the entire repository.

    Starts with files in the canonical *adr_dir*, then walks the rest of
    the repository tree to find ADR files that live outside that directory.
    Directories listed in ``_SCAN_SKIP_DIRS`` are pruned for performance.
    """
    seen_paths: set[Path] = set()
    ordered: list[Path] = []

    resolved_adr_dir = adr_dir.resolve() if adr_dir.is_dir() else None

    # 1. Canonical ADR directory (non-recursive, preserves existing behaviour)
    if resolved_adr_dir is not None:
        for fp in sorted(adr_dir.glob("*.md")):
            resolved = fp.resolve()
            if resolved not in seen_paths:
                seen_paths.add(resolved)
                ordered.append(fp)

    # 2. Walk the whole repo for additional ADR files
    for dirpath, dirnames, filenames in os.walk(repo_root):
        # Prune skipped directories (modifying dirnames in-place)
        dirnames[:] = [
            d for d in dirnames if d not in _SCAN_SKIP_DIRS
        ]
        dp = Path(dirpath)
        # Skip the canonical ADR directory — already processed above
        if resolved_adr_dir is not None and dp.resolve() == resolved_adr_dir:
            continue
        for fn in sorted(filenames):
            if not fn.endswith(".md"):
                continue
            fp = dp / fn
            resolved = fp.resolve()
            if resolved in seen_paths:
                continue
            seen_paths.add(resolved)
            ordered.append(fp)

    return ordered


# ---------------------------------------------------------------------------
# Markdown / YAML parsing
# ---------------------------------------------------------------------------

_FRONTMATTER_RE = re.compile(r"\A---\n(.*?\n)---\n", re.DOTALL)
_ADR_ID_RE = re.compile(r"^\s*adr_id:\s*(\S+)", re.MULTILINE)

# Regex to extract a numeric ADR identifier from a filename.
# Matches patterns like: ADR-001-title.md, adr-001-title.md, 0042-title.md
_FILENAME_ADR_ID_RE = re.compile(
    r"^(?:adr[_-]?)?(\d+)(?:-[a-zA-Z].*)?\.md$", re.IGNORECASE,
)


def _parse_frontmatter(text: str) -> dict[str, Any]:
    """Extract YAML frontmatter from a markdown string."""
    m = _FRONTMATTER_RE.match(text)
    if not m:
        return {}
    raw = m.group(1)
    if yaml is not None:
        return yaml.safe_load(raw) or {}
    # Minimal fallback when PyYAML is absent — only extracts flat top-level
    # key: value pairs.  Nested structures (ejs, actors, context) will not
    # be parsed correctly, so adr-db features that depend on them will
    # degrade.  Install PyYAML for full functionality.
    result: dict[str, Any] = {}
    for line in raw.splitlines():
        if ":" in line and not line.startswith(" "):
            key, _, value = line.partition(":")
            result[key.strip()] = value.strip()
    return result


def _extract_section(text: str, heading: str) -> str:
    """Extract content under a markdown ## or # heading, up to the next heading of same or higher level."""
    # Match heading at level 1 or 2
    pattern = re.compile(
        rf"^(#{{1,2}})\s+{re.escape(heading)}\s*\n(.*?)(?=\n#{{1,2}}\s|\Z)",
        re.MULTILINE | re.DOTALL,
    )
    m = pattern.search(text)
    return m.group(2).strip() if m else ""


def parse_adr_file(filepath: Path) -> dict[str, Any] | None:
    """Parse a single ADR markdown file into a dict of metadata + sections.

    Supports two formats:
    1. **EJS format** — YAML frontmatter with an ``ejs:`` block containing
       ``adr_id``, ``title``, ``status``, etc.
    2. **Plain / Nygard format** — No frontmatter; uses markdown headings
       (``# Title``, ``## Status``, ``## Context``, ``## Decision``, …).

    The EJS format is tried first.  If the file lacks EJS frontmatter the
    plain parser is used as a fallback so that conventional ADRs created by
    sub-agents or other tooling are still indexed.
    """
    text = filepath.read_text(encoding="utf-8")

    fm = _parse_frontmatter(text)

    # --- EJS-format path ---
    if fm:
        ejs = fm.get("ejs", {})
        if isinstance(ejs, dict) and ejs:
            # Extract adr_id from raw frontmatter text to avoid YAML octal
            # interpretation (e.g. 0042 parsed as decimal 34).
            fm_match = _FRONTMATTER_RE.match(text)
            raw_fm = fm_match.group(1) if fm_match else ""
            id_match = _ADR_ID_RE.search(raw_fm)
            adr_id = id_match.group(1) if id_match else ""
            if not adr_id or adr_id == "XXXX":
                return None  # Skip template

            actors = fm.get("actors", {}) or {}
            ctx = fm.get("context", {}) or {}

            return {
                "adr_id": adr_id,
                "title": ejs.get("title", ""),
                "date": str(ejs.get("date", "")),
                "status": ejs.get("status", ""),
                "session_id": ejs.get("session_id", ""),
                "session_journey": ejs.get("session_journey", ""),
                "actors_humans": json.dumps(actors.get("humans", [])),
                "actors_agents": json.dumps(actors.get("agents", [])),
                "context_repo": ctx.get("repo", ""),
                "context_branch": ctx.get("branch", ""),
                "decision": _extract_section(text, "Decision"),
                "context_section": _extract_section(text, "Context"),
                "rationale": _extract_section(text, "Rationale"),
                "consequences": _extract_section(text, "Consequences"),
                "key_learnings": _extract_section(text, "Key Learnings"),
                "agent_guidance": _extract_section(text, "Agent Guidance"),
                "file_path": str(filepath.relative_to(_repo_root())),
            }

    # --- Plain / Nygard-format fallback ---
    return _parse_plain_adr(text, filepath)


def _parse_plain_adr(text: str, filepath: Path) -> dict[str, Any] | None:
    """Parse a plain (Nygard-style) ADR that has no YAML frontmatter.

    Expected structure::

        # <Number>. <Title>    — or —    # ADR-<Number>: <Title>
        ## Status
        <status text>
        ## Context
        ...
        ## Decision
        ...
        ## Consequences
        ...

    The ADR identifier is derived from the first ``#`` heading or, failing
    that, from the filename (e.g. ``ADR-001-some-title.md`` → ``ADR-001``).
    Files that cannot be identified as ADRs are skipped (returns ``None``).
    """
    # Try to extract title + id from the first H1 heading
    title_match = re.search(r"^#\s+(.+)", text, re.MULTILINE)
    if not title_match:
        return None  # No heading — not an ADR

    raw_title = title_match.group(1).strip()

    # Derive adr_id from the heading text.
    # Patterns: "1. Title", "ADR-001: Title", "ADR 001 - Title", "001 - Title"
    id_from_heading = re.match(
        r"(?:ADR[_\s-]*)?(\d+)[\s.:)\-—]+\s*(.*)",
        raw_title,
        re.IGNORECASE,
    )

    if id_from_heading:
        adr_id = f"ADR-{id_from_heading.group(1).zfill(3)}"
        title = id_from_heading.group(2).strip() or raw_title
    else:
        # Fall back to filename for the id
        fn_match = _FILENAME_ADR_ID_RE.match(filepath.name)
        if not fn_match:
            return None  # Cannot determine an id — skip
        adr_id = f"ADR-{fn_match.group(1).zfill(3)}"
        title = raw_title

    # Skip obvious template files
    lower_title = title.lower()
    if "template" in lower_title and adr_id == "ADR-000":
        return None

    # A valid plain ADR must have at least a Decision or Context section
    decision = _extract_section(text, "Decision")
    context_section = _extract_section(text, "Context")
    if not decision and not context_section:
        return None

    # Extract Status from the ## Status section
    status_raw = _extract_section(text, "Status")
    status_text = status_raw.split("\n", 1)[0].strip() if status_raw else ""

    # Extract date from ## Date section if present
    date_raw = _extract_section(text, "Date")
    date_text = date_raw.split("\n", 1)[0].strip() if date_raw else ""

    return {
        "adr_id": adr_id,
        "title": title,
        "date": date_text,
        "status": status_text.lower() if status_text else "",
        "session_id": "",
        "session_journey": "",
        "actors_humans": json.dumps([]),
        "actors_agents": json.dumps([]),
        "context_repo": "",
        "context_branch": "",
        "decision": decision,
        "context_section": context_section,
        "rationale": _extract_section(text, "Rationale"),
        "consequences": _extract_section(text, "Consequences"),
        "key_learnings": _extract_section(text, "Key Learnings"),
        "agent_guidance": _extract_section(text, "Agent Guidance"),
        "file_path": str(filepath.relative_to(_repo_root())),
    }


# Session ID regex for raw-text extraction from journey metadata
_SESSION_ID_RE = re.compile(r"^\s*session_id:\s*(\S+)", re.MULTILINE)


def _parse_journey_metadata(text: str) -> dict[str, str]:
    """Extract flat key: value metadata from the top of a journey file.

    Journey files may use YAML frontmatter (``---`` delimiters) or plain
    ``key: value`` lines before the first ``#`` heading.  This helper
    normalises both formats into a flat dict of string values.
    """
    fm_match = _FRONTMATTER_RE.match(text)
    if fm_match:
        raw = fm_match.group(1)
    else:
        # Plain key: value lines before the first heading
        first_heading = re.search(r"^#", text, re.MULTILINE)
        raw = text[: first_heading.start()] if first_heading else text

    if yaml is not None and fm_match:
        parsed = yaml.safe_load(raw) or {}
        # Flatten to strings for uniform handling
        return {k: str(v) if not isinstance(v, str) else v for k, v in parsed.items()}

    result: dict[str, str] = {}
    for line in raw.splitlines():
        if ":" in line and not line.startswith(" ") and not line.startswith("#"):
            key, _, value = line.partition(":")
            result[key.strip()] = value.strip()
    return result


def parse_journey_file(filepath: Path) -> dict[str, Any] | None:
    """Parse a single Session Journey markdown file into a dict."""
    text = filepath.read_text(encoding="utf-8")

    meta = _parse_journey_metadata(text)
    session_id = meta.get("session_id", "").strip()
    if not session_id:
        return None  # Skip template or invalid files

    return {
        "session_id": session_id,
        "author": meta.get("author", ""),
        "date": meta.get("date", ""),
        "repo": meta.get("repo", ""),
        "branch": meta.get("branch", ""),
        "agents_involved": meta.get("agents_involved", ""),
        "decision_detected": str(meta.get("decision_detected", "false")).lower(),
        "adr_links": meta.get("adr_links", ""),
        "tags": meta.get("tags", ""),
        "problem_intent": _extract_section(text, "Problem / Intent"),
        "interaction_summary": _extract_section(text, "Interaction Summary")
        or _extract_section(text, "Interaction Summary (Required)"),
        "decisions_made": _extract_section(text, "Decisions Made"),
        "key_learnings": _extract_section(text, "Key Learnings"),
        "future_agent_guidance": _extract_section(text, "Future Agent Guidance"),
        "file_path": str(filepath.relative_to(_repo_root())),
    }


# ---------------------------------------------------------------------------
# Database layer
# ---------------------------------------------------------------------------

_SCHEMA = """
CREATE TABLE IF NOT EXISTS adrs (
    adr_id          TEXT PRIMARY KEY,
    title           TEXT NOT NULL,
    date            TEXT,
    status          TEXT,
    session_id      TEXT,
    session_journey TEXT,
    actors_humans   TEXT,
    actors_agents   TEXT,
    context_repo    TEXT,
    context_branch  TEXT,
    decision        TEXT,
    context_section TEXT,
    rationale       TEXT,
    consequences    TEXT,
    key_learnings   TEXT,
    agent_guidance  TEXT,
    file_path       TEXT,
    last_synced     TEXT
);

CREATE VIRTUAL TABLE IF NOT EXISTS adrs_fts USING fts5(
    adr_id,
    title,
    decision,
    context_section,
    rationale,
    consequences,
    key_learnings,
    agent_guidance,
    content='adrs',
    content_rowid='rowid'
);

CREATE TRIGGER IF NOT EXISTS adrs_ai AFTER INSERT ON adrs BEGIN
    INSERT INTO adrs_fts(rowid, adr_id, title, decision, context_section,
                         rationale, consequences, key_learnings, agent_guidance)
    VALUES (new.rowid, new.adr_id, new.title, new.decision, new.context_section,
            new.rationale, new.consequences, new.key_learnings, new.agent_guidance);
END;

CREATE TRIGGER IF NOT EXISTS adrs_ad AFTER DELETE ON adrs BEGIN
    INSERT INTO adrs_fts(adrs_fts, rowid, adr_id, title, decision, context_section,
                         rationale, consequences, key_learnings, agent_guidance)
    VALUES ('delete', old.rowid, old.adr_id, old.title, old.decision, old.context_section,
            old.rationale, old.consequences, old.key_learnings, old.agent_guidance);
END;

CREATE TRIGGER IF NOT EXISTS adrs_au AFTER UPDATE ON adrs BEGIN
    INSERT INTO adrs_fts(adrs_fts, rowid, adr_id, title, decision, context_section,
                         rationale, consequences, key_learnings, agent_guidance)
    VALUES ('delete', old.rowid, old.adr_id, old.title, old.decision, old.context_section,
            old.rationale, old.consequences, old.key_learnings, old.agent_guidance);
    INSERT INTO adrs_fts(rowid, adr_id, title, decision, context_section,
                         rationale, consequences, key_learnings, agent_guidance)
    VALUES (new.rowid, new.adr_id, new.title, new.decision, new.context_section,
            new.rationale, new.consequences, new.key_learnings, new.agent_guidance);
END;
"""

_UPSERT = """
INSERT INTO adrs (
    adr_id, title, date, status, session_id, session_journey,
    actors_humans, actors_agents, context_repo, context_branch,
    decision, context_section, rationale, consequences,
    key_learnings, agent_guidance, file_path, last_synced
) VALUES (
    :adr_id, :title, :date, :status, :session_id, :session_journey,
    :actors_humans, :actors_agents, :context_repo, :context_branch,
    :decision, :context_section, :rationale, :consequences,
    :key_learnings, :agent_guidance, :file_path, :last_synced
)
ON CONFLICT(adr_id) DO UPDATE SET
    title           = excluded.title,
    date            = excluded.date,
    status          = excluded.status,
    session_id      = excluded.session_id,
    session_journey = excluded.session_journey,
    actors_humans   = excluded.actors_humans,
    actors_agents   = excluded.actors_agents,
    context_repo    = excluded.context_repo,
    context_branch  = excluded.context_branch,
    decision        = excluded.decision,
    context_section = excluded.context_section,
    rationale       = excluded.rationale,
    consequences    = excluded.consequences,
    key_learnings   = excluded.key_learnings,
    agent_guidance  = excluded.agent_guidance,
    file_path       = excluded.file_path,
    last_synced     = excluded.last_synced;
"""

_JOURNEYS_SCHEMA = """
CREATE TABLE IF NOT EXISTS journeys (
    session_id              TEXT PRIMARY KEY,
    author                  TEXT,
    date                    TEXT,
    repo                    TEXT,
    branch                  TEXT,
    agents_involved         TEXT,
    decision_detected       TEXT,
    adr_links               TEXT,
    tags                    TEXT,
    problem_intent          TEXT,
    interaction_summary     TEXT,
    decisions_made          TEXT,
    key_learnings           TEXT,
    future_agent_guidance   TEXT,
    file_path               TEXT,
    last_synced             TEXT
);

CREATE VIRTUAL TABLE IF NOT EXISTS journeys_fts USING fts5(
    session_id,
    problem_intent,
    interaction_summary,
    decisions_made,
    key_learnings,
    future_agent_guidance,
    content='journeys',
    content_rowid='rowid'
);

CREATE TRIGGER IF NOT EXISTS journeys_ai AFTER INSERT ON journeys BEGIN
    INSERT INTO journeys_fts(rowid, session_id, problem_intent, interaction_summary,
                             decisions_made, key_learnings, future_agent_guidance)
    VALUES (new.rowid, new.session_id, new.problem_intent, new.interaction_summary,
            new.decisions_made, new.key_learnings, new.future_agent_guidance);
END;

CREATE TRIGGER IF NOT EXISTS journeys_ad AFTER DELETE ON journeys BEGIN
    INSERT INTO journeys_fts(journeys_fts, rowid, session_id, problem_intent, interaction_summary,
                             decisions_made, key_learnings, future_agent_guidance)
    VALUES ('delete', old.rowid, old.session_id, old.problem_intent, old.interaction_summary,
            old.decisions_made, old.key_learnings, old.future_agent_guidance);
END;

CREATE TRIGGER IF NOT EXISTS journeys_au AFTER UPDATE ON journeys BEGIN
    INSERT INTO journeys_fts(journeys_fts, rowid, session_id, problem_intent, interaction_summary,
                             decisions_made, key_learnings, future_agent_guidance)
    VALUES ('delete', old.rowid, old.session_id, old.problem_intent, old.interaction_summary,
            old.decisions_made, old.key_learnings, old.future_agent_guidance);
    INSERT INTO journeys_fts(rowid, session_id, problem_intent, interaction_summary,
                             decisions_made, key_learnings, future_agent_guidance)
    VALUES (new.rowid, new.session_id, new.problem_intent, new.interaction_summary,
            new.decisions_made, new.key_learnings, new.future_agent_guidance);
END;
"""

_JOURNEY_UPSERT = """
INSERT INTO journeys (
    session_id, author, date, repo, branch, agents_involved,
    decision_detected, adr_links, tags, problem_intent,
    interaction_summary, decisions_made, key_learnings,
    future_agent_guidance, file_path, last_synced
) VALUES (
    :session_id, :author, :date, :repo, :branch, :agents_involved,
    :decision_detected, :adr_links, :tags, :problem_intent,
    :interaction_summary, :decisions_made, :key_learnings,
    :future_agent_guidance, :file_path, :last_synced
)
ON CONFLICT(session_id) DO UPDATE SET
    author                  = excluded.author,
    date                    = excluded.date,
    repo                    = excluded.repo,
    branch                  = excluded.branch,
    agents_involved         = excluded.agents_involved,
    decision_detected       = excluded.decision_detected,
    adr_links               = excluded.adr_links,
    tags                    = excluded.tags,
    problem_intent          = excluded.problem_intent,
    interaction_summary     = excluded.interaction_summary,
    decisions_made          = excluded.decisions_made,
    key_learnings           = excluded.key_learnings,
    future_agent_guidance   = excluded.future_agent_guidance,
    file_path               = excluded.file_path,
    last_synced             = excluded.last_synced;
"""


def init_db(db_path: Path) -> sqlite3.Connection:
    """Open (or create) the database and ensure schema exists."""
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.executescript(_SCHEMA)
    conn.executescript(_JOURNEYS_SCHEMA)
    return conn


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def cmd_sync(conn: sqlite3.Connection, adr_dir: Path,
             journey_dir: Path | None = None) -> int:
    """Parse ADR (and optionally journey) files and upsert into the database."""
    rc = _sync_adrs(conn, adr_dir)
    if journey_dir is not None and journey_dir.is_dir():
        rc2 = _sync_journeys(conn, journey_dir)
        if rc == 0:
            rc = rc2
    return rc


def _sync_adrs(conn: sqlite3.Connection, adr_dir: Path) -> int:
    """Sync ADR markdown files into the database.

    Scans the canonical *adr_dir* **and** the entire repository tree so
    that ADRs placed in non-standard locations are also discovered and
    indexed.
    """
    if not adr_dir.is_dir():
        print(f"ADR directory not found: {adr_dir}", file=sys.stderr)
        return 1

    repo_root = _repo_root()
    all_files = _discover_adr_files(repo_root, adr_dir)

    now = datetime.now(timezone.utc).isoformat()
    count = 0
    extra = 0
    resolved_adr_dir = adr_dir.resolve()
    disk_ids: set[str] = set()
    for fp in all_files:
        record = parse_adr_file(fp)
        if record is None:
            continue
        disk_ids.add(record["adr_id"])
        record["last_synced"] = now
        conn.execute(_UPSERT, record)
        count += 1
        # Track ADRs found outside the canonical directory
        try:
            fp.resolve().relative_to(resolved_adr_dir)
        except ValueError:
            extra += 1

    # Remove ADRs that no longer exist on disk
    db_ids = {row["adr_id"] for row in conn.execute("SELECT adr_id FROM adrs")}
    stale = db_ids - disk_ids
    for sid in stale:
        conn.execute("DELETE FROM adrs WHERE adr_id = ?", (sid,))

    conn.commit()
    msg = f"Synced {count} ADR(s). Removed {len(stale)} stale record(s)."
    if extra:
        msg += f" ({extra} found outside {adr_dir})"
    print(msg)
    return 0


def _sync_journeys(conn: sqlite3.Connection, journey_dir: Path) -> int:
    """Sync Session Journey markdown files into the database."""
    if not journey_dir.is_dir():
        print(f"Journey directory not found: {journey_dir}", file=sys.stderr)
        return 1

    now = datetime.now(timezone.utc).isoformat()
    count = 0
    disk_ids: set[str] = set()
    for fp in sorted(journey_dir.rglob("*.md")):
        # Skip template files
        if "_templates" in fp.parts:
            continue
        record = parse_journey_file(fp)
        if record is None:
            continue
        disk_ids.add(record["session_id"])
        record["last_synced"] = now
        conn.execute(_JOURNEY_UPSERT, record)
        count += 1

    # Remove journeys that no longer exist on disk
    db_ids = {row["session_id"] for row in conn.execute("SELECT session_id FROM journeys")}
    stale = db_ids - disk_ids
    for sid in stale:
        conn.execute("DELETE FROM journeys WHERE session_id = ?", (sid,))

    conn.commit()
    print(f"Synced {count} journey(s). Removed {len(stale)} stale record(s).")
    return 0


def cmd_list(conn: sqlite3.Connection) -> int:
    """List all ADRs in compact format."""
    rows = conn.execute(
        "SELECT adr_id, title, date, status FROM adrs ORDER BY adr_id"
    ).fetchall()
    if not rows:
        print("No ADRs in database. Run 'sync' first.")
        return 0
    for r in rows:
        print(f"[{r['adr_id']}] {r['title']}  ({r['status']}, {r['date']})")
    return 0


def cmd_get(conn: sqlite3.Connection, adr_id: str) -> int:
    """Show full details for a single ADR."""
    row = conn.execute("SELECT * FROM adrs WHERE adr_id = ?", (adr_id,)).fetchone()
    if not row:
        print(f"ADR {adr_id} not found.", file=sys.stderr)
        return 1
    for key in row.keys():
        val = row[key]
        if key in ("actors_humans", "actors_agents"):
            val = json.dumps(json.loads(val), indent=2)
        print(f"--- {key} ---")
        print(val)
        print()
    return 0


def _sanitise_fts_query(query: str) -> str:
    """Quote bare words that contain FTS5 special characters (e.g. hyphens).

    FTS5 interprets ``-`` as the NOT operator.  This helper wraps tokens
    that contain such characters in double-quotes so that a plain search
    like ``sub-agent`` works as expected without requiring the caller to
    know FTS5 syntax.
    """
    if '"' in query:
        return query  # Already contains explicit quoting
    tokens = query.split()
    # Quote bare words that contain FTS5 special chars (hyphens, dots, etc.)
    # so that e.g. "sub-agent" is not interpreted as "sub NOT agent".
    out: list[str] = []
    for t in tokens:
        if re.search(r"[^\w*]", t):
            out.append(f'"{t}"')
        else:
            out.append(t)
    return " ".join(out)


def cmd_search(conn: sqlite3.Connection, query: str) -> int:
    """Full-text search across ADR and journey content."""
    fts_query = _sanitise_fts_query(query)
    adr_rows = conn.execute(
        """
        SELECT a.adr_id, a.title, a.status, a.date,
               snippet(adrs_fts, 2, '>>>', '<<<', '...', 32) AS snippet
        FROM adrs_fts
        JOIN adrs a ON a.rowid = adrs_fts.rowid
        WHERE adrs_fts MATCH ?
        ORDER BY rank
        """,
        (fts_query,),
    ).fetchall()
    journey_rows = conn.execute(
        """
        SELECT j.session_id, j.date, j.decision_detected,
               snippet(journeys_fts, 1, '>>>', '<<<', '...', 32) AS snippet
        FROM journeys_fts
        JOIN journeys j ON j.rowid = journeys_fts.rowid
        WHERE journeys_fts MATCH ?
        ORDER BY rank
        """,
        (fts_query,),
    ).fetchall()
    if not adr_rows and not journey_rows:
        print("No results.")
        return 0
    for r in adr_rows:
        print(f"[ADR {r['adr_id']}] {r['title']}  ({r['status']}, {r['date']})")
        if r["snippet"]:
            print(f"  …{r['snippet']}…")
        print()
    for r in journey_rows:
        print(f"[Journey {r['session_id']}]  (decision: {r['decision_detected']}, {r['date']})")
        if r["snippet"]:
            print(f"  …{r['snippet']}…")
        print()
    return 0


def cmd_summary(conn: sqlite3.Connection) -> int:
    """Produce a compact, agent-friendly summary of all ADRs."""
    rows = conn.execute(
        "SELECT adr_id, title, date, status, decision, key_learnings, agent_guidance "
        "FROM adrs ORDER BY adr_id"
    ).fetchall()
    if not rows:
        print("No ADRs in database. Run 'sync' first.")
        return 0

    parts: list[str] = []
    for r in rows:
        decision_short = (r["decision"] or "")[:300]
        learnings_short = (r["key_learnings"] or "")[:200]
        guidance_short = (r["agent_guidance"] or "")[:200]
        parts.append(
            f"### ADR {r['adr_id']}: {r['title']}\n"
            f"Status: {r['status']} | Date: {r['date']}\n"
            f"Decision: {decision_short}\n"
            f"Learnings: {learnings_short}\n"
            f"Agent Guidance: {guidance_short}\n"
        )
    print("\n".join(parts))
    return 0


def cmd_list_journeys(conn: sqlite3.Connection) -> int:
    """List all journeys in compact format."""
    rows = conn.execute(
        "SELECT session_id, date, decision_detected, repo, branch FROM journeys ORDER BY date, session_id"
    ).fetchall()
    if not rows:
        print("No journeys in database. Run 'sync' first.")
        return 0
    for r in rows:
        print(f"[{r['session_id']}]  (decision: {r['decision_detected']}, {r['date']}, {r['repo']})")
    return 0


def cmd_get_journey(conn: sqlite3.Connection, session_id: str) -> int:
    """Show full details for a single journey."""
    row = conn.execute("SELECT * FROM journeys WHERE session_id = ?", (session_id,)).fetchone()
    if not row:
        print(f"Journey {session_id} not found.", file=sys.stderr)
        return 1
    for key in row.keys():
        print(f"--- {key} ---")
        print(row[key])
        print()
    return 0


def cmd_summary_journeys(conn: sqlite3.Connection) -> int:
    """Produce a compact, agent-friendly summary of all journeys."""
    rows = conn.execute(
        "SELECT session_id, date, decision_detected, problem_intent, "
        "key_learnings, future_agent_guidance FROM journeys ORDER BY date, session_id"
    ).fetchall()
    if not rows:
        print("No journeys in database. Run 'sync' first.")
        return 0

    parts: list[str] = []
    for r in rows:
        intent_short = (r["problem_intent"] or "")[:300]
        learnings_short = (r["key_learnings"] or "")[:200]
        guidance_short = (r["future_agent_guidance"] or "")[:200]
        parts.append(
            f"### {r['session_id']}\n"
            f"Date: {r['date']} | Decision: {r['decision_detected']}\n"
            f"Intent: {intent_short}\n"
            f"Learnings: {learnings_short}\n"
            f"Agent Guidance: {guidance_short}\n"
        )
    print("\n".join(parts))
    return 0


# ---------------------------------------------------------------------------
# CLI entry-point
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="EJS Database — SQLite index for ADRs and Session Journeys",
    )
    parser.add_argument(
        "--db",
        type=Path,
        default=None,
        help="Path to SQLite database (default: <repo>/.ejs.db)",
    )
    parser.add_argument(
        "--adr-dir",
        type=Path,
        default=None,
        help="Path to ADR markdown directory (default: <repo>/ejs-docs/adr)",
    )
    parser.add_argument(
        "--journey-dir",
        type=Path,
        default=None,
        help="Path to journey markdown directory (default: <repo>/ejs-docs/journey)",
    )
    sub = parser.add_subparsers(dest="command")

    sub.add_parser("sync", help="Parse ADR and journey files and update the database")
    sub.add_parser("list", help="List all ADRs")

    p_get = sub.add_parser("get", help="Show full details for an ADR")
    p_get.add_argument("adr_id", help="ADR identifier (e.g. 0010)")

    p_search = sub.add_parser("search", help="Full-text search across ADRs and journeys")
    p_search.add_argument("query", help="Search query")

    sub.add_parser("summary", help="Agent-friendly compact summary of ADRs")

    sub.add_parser("list-journeys", help="List all journeys")

    p_get_j = sub.add_parser("get-journey", help="Show full details for a journey")
    p_get_j.add_argument("session_id", help="Session identifier (e.g. ejs-session-2026-02-10-01)")

    sub.add_parser("summary-journeys", help="Agent-friendly compact summary of journeys")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        return 1

    db_path = args.db or _default_db_path()
    adr_dir = args.adr_dir or _default_adr_dir()
    journey_dir = args.journey_dir or _default_journey_dir()

    conn = init_db(db_path)
    try:
        if args.command == "sync":
            return cmd_sync(conn, adr_dir, journey_dir)
        if args.command == "list":
            return cmd_list(conn)
        if args.command == "get":
            return cmd_get(conn, args.adr_id)
        if args.command == "search":
            return cmd_search(conn, args.query)
        if args.command == "summary":
            return cmd_summary(conn)
        if args.command == "list-journeys":
            return cmd_list_journeys(conn)
        if args.command == "get-journey":
            return cmd_get_journey(conn, args.session_id)
        if args.command == "summary-journeys":
            return cmd_summary_journeys(conn)
        parser.print_help()
        return 1
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
