---
name: ejs-session-wrapup
description: >
  Finalize an Engineering Journey System (EJS) session by completing all
  journey sections, populating machine extracts, evaluating the ADR decision
  rubric, and optionally creating an ADR document.
---

# EJS Session Wrap-Up

Use this skill when a session is ending — for example when the user says
"wrap up", "finalize session", "end session", "ship it", "commit this", or
"commit and push".

This skill also applies for **context-threshold checkpoints** — proactive
mid-session saves that prevent documentation loss if context runs out. See
the Checkpoint vs. Full Finalization section below for the differences.

## Steps

1. **Review the Session Journey for completeness**
   - Read the current journey file at `ejs-docs/journey/YYYY/ejs-session-YYYY-MM-DD-<seq>.md`
   - Identify any sections that are incomplete or missing context.

2. **Finalize all sections**
   Complete each section with coherent summaries:
   - **Interaction Summary** — ensure all key exchanges are documented
   - **Agent Collaboration Summary** — which agents participated and their contributions
   - **Sub-Agent Contributions** — if sub-agents were involved, ensure their decisions and handoffs are captured
   - **Agent Influence** — suggestions adopted vs. rejected, human overrides
   - **Experiments / Evidence** — what was tried and what happened
   - **Iteration Log** — pivots, reversals, or refinements
   - **Decisions Made** — all decisions with reason and impact
   - **Key Learnings** — technical, prompting, and tooling insights
   - **If Repeating This Work** — do this, avoid this, watch out for
   - **Future Agent Guidance** — prefer/avoid patterns for future agents

3. **Populate machine extracts**
   Fill in the `## MACHINE EXTRACTS` section with structured summaries:
   - `INTERACTION_EXTRACT` — compact summary of the collaboration trail
   - `DECISIONS_EXTRACT` — list of decisions with rationale
   - `LEARNING_EXTRACT` — transferable insights
   - `AGENT_GUIDANCE_EXTRACT` — guidance for future agents
   - `SUB_AGENT_EXTRACT` — sub-agent contributions (if applicable)

4. **Evaluate the ADR decision rubric**
   Create an ADR only if **at least one** of these criteria is met:
   - Introduces or changes a system boundary (service, datastore, topology)
   - Changes a public contract (API, schema, protocol)
   - Alters security, privacy, or compliance posture
   - Requires choosing among credible alternatives with meaningful trade-offs
   - Has long-lived or hard-to-reverse consequences
   - Changes engineering process or workflow for future work

5. **Update `decision_detected` field**
   - Set to `true` if an ADR is warranted, `false` otherwise.

6. **Create ADR if warranted**
   - Use template: `ejs-docs/adr/0000-adr-template.md`
   - Save to: `ejs-docs/adr/NNNN-<kebab-title>.md` (next available number)
   - Link the ADR back to the session journey and vice versa via `adr_links`.

7. **Confirm finalization**
   - Inform the user: `"Session finalized: ejs-session-YYYY-MM-DD-<seq>"`
   - If an ADR was created, mention it: `"ADR NNNN created: <title>"`

## Contextual References

- ADR template: `ejs-docs/adr/0000-adr-template.md`
- Lifecycle patterns: `ejs-docs/session-lifecycle-patterns.md`
- Database tool: `scripts/adr-db.py`

## Key Principle

By session end, most of the journey should already be populated from
continuous updates. Finalization is a quick **review and completion** step,
not a full reconstruction effort.

## Checkpoint vs. Full Finalization

| Aspect | Checkpoint (mid-session) | Full Finalization (session end) |
|--------|--------------------------|-------------------------------|
| **Trigger** | Context getting large, 3+ unsaved interactions, before heavy operations, 5+ exchanges since last save | User signals session end |
| **Sections updated** | Interaction Summary, Decisions Made, Experiments, Iteration Log, Key Learnings | All sections reviewed and completed |
| **Machine extracts** | Not populated | Populated in full |
| **ADR rubric** | Not evaluated | Evaluated; ADR created if warranted |
| **Goal** | Preserve work-in-progress against context loss | Produce a coherent, complete record |

### When to Perform a Checkpoint
- **3+ unsaved interactions** have accumulated since the last save (an interaction is one human prompt and the corresponding agent response)
- A **significant decision** has been made but not yet written to the journey
- A **large, context-intensive operation** is about to start
- **5+ exchanges** have occurred since the last journey file save
- Substantial work completed but user has not signalled session end