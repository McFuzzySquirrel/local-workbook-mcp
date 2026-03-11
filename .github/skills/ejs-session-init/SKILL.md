---
name: ejs-session-init
description: >
  Initialize an Engineering Journey System (EJS) session by creating a
  Session Journey file, refreshing the ADR/journey database index, and
  setting up metadata for continuous recording throughout the session.
---

# EJS Session Initialization

Use this skill when a session is starting — for example when the user says
"initialize session", "start session", "create session journey", or begins
work on a new task, feature, or bug fix.

## Steps

1. **Refresh the EJS database index**
   ```bash
   python scripts/adr-db.py sync
   ```
   This ensures past ADRs and session journeys are queryable for context.

2. **Determine the session ID**
   - Format: `ejs-session-YYYY-MM-DD-<seq>` (2-digit daily sequence)
   - Check `ejs-docs/journey/YYYY/` for existing sessions today and increment the sequence number.

3. **Create the Session Journey file**
   - Copy the template from `ejs-docs/journey/_templates/journey-template.md`
   - Save to `ejs-docs/journey/YYYY/ejs-session-YYYY-MM-DD-<seq>.md`

4. **Populate initial metadata**
   Fill in the frontmatter fields:
   - `session_id` — the generated session ID
   - `author` — the human user (if known)
   - `date` — today's date (YYYY-MM-DD)
   - `repo` — current repository name
   - `branch` — current git branch
   - `agents_involved` — list the active agents (e.g., `[copilot]`)
   - `decision_detected: false` — initial value; updated at session end
   - `adr_links: []` — populated if an ADR is created later

5. **Capture initial Problem / Intent**
   - Write the user's stated goal or task description into the **Problem / Intent** section.

6. **Confirm initialization**
   - Inform the user: `"Session initialized: ejs-session-YYYY-MM-DD-<seq>"`

## Contextual References

- Session Journey template: `ejs-docs/journey/_templates/journey-template.md`
- ADR template: `ejs-docs/adr/0000-adr-template.md`
- Lifecycle patterns: `ejs-docs/session-lifecycle-patterns.md`
- Database tool: `scripts/adr-db.py`

## Key Principle

Initialize early, capture context while it's fresh. A well-initialized
session reduces the burden at session end and produces higher-quality
documentation.
