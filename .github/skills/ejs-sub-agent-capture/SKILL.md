---
name: ejs-sub-agent-capture
description: >
  Capture sub-agent decisions, rationale, and inter-agent handoffs during
  multi-agent Engineering Journey System (EJS) workflows. Ensures that
  delegated work is properly attributed and traceable.
---

# EJS Sub-Agent Capture

Use this skill during multi-agent workflows — when a main agent delegates
tasks to sub-agents (e.g., explore, task, general-purpose, or custom agents)
and their decisions and outputs need to be captured in the Session Journey.

## When Delegating to a Sub-Agent

1. **Record the delegation** in the **Interaction Summary**:
   ```
   - Human: <request that triggered delegation>
     - Agent [main-agent]: Delegated [task] to sub-agent [sub-agent-name]
     - Context provided: <what constraints or prior work was shared>
   ```

2. **Provide context from prior sub-agent work** when relevant:
   - Pass earlier sub-agent findings to subsequent sub-agents
   - Example: code review findings → testing agent, test results → documentation agent

3. **Instruct sub-agents** to follow EJS conventions:
   - Use the ADR template at `ejs-docs/adr/0000-adr-template.md`
   - Place ADRs under `ejs-docs/adr/`
   - Do not let sub-agents create conflicting templates or instruction files

## When a Sub-Agent Completes

4. **Capture the sub-agent's contribution** in the **Sub-Agent Contributions** section:

   ```markdown
   ## Sub-Agent: <agent-name / agent-type>
   - **Task delegated:** What was the sub-agent asked to do?
   - **Decisions made:** What did the sub-agent decide (approach, trade-offs)?
   - **Alternatives considered:** What other approaches did the sub-agent evaluate?
   - **Outcome:** What was the result?
   - **Handoff to other agents:** Did this sub-agent's output feed into another's work?
   ```

5. **Record inter-agent handoff chains** when one sub-agent's output informs another:

   ```
   SA1 (Code Review) → found missing input validation
     ↓ handoff
   SA2 (Testing) → added edge-case tests based on SA1 findings
     ↓ handoff
   SA3 (Documentation) → added validation examples based on SA2 test cases
   ```

   Document:
   - The dependency: which sub-agent's output became another's input
   - Any disagreements between sub-agents and how conflicts were resolved
   - Which agent ultimately influenced the final decision

## At Session Finalization

6. **Populate the `SUB_AGENT_EXTRACT`** in the machine extracts section with:
   - List of sub-agents involved (by name/type)
   - Summary of each sub-agent's decisions and rationale
   - Inter-agent handoff chain documentation
   - Any disagreements or conflicts resolved

## Example

```markdown
## Sub-Agent: Code Review (explore agent)
- **Task delegated:** Review authentication refactor for security and code quality
- **Decisions made:** Flagged missing input validation on JWT claims (chose depth-first security review over breadth-first style review)
- **Alternatives considered:** Could have focused on code style first, but prioritized security given auth context
- **Outcome:** 3 security findings, 1 code quality suggestion
- **Handoff to other agents:** Findings passed to Testing agent as context for edge-case test generation

## Sub-Agent: Testing (task agent)
- **Task delegated:** Run and extend test suite for auth changes, with Code Review findings as context
- **Decisions made:** Added 3 new edge-case tests targeting validation gaps found by Code Review
- **Alternatives considered:** Could have run existing tests only, but SA1 findings warranted new test cases
- **Outcome:** 2 test failures found and fixed, 3 new tests added
- **Handoff to other agents:** Test results and new test cases passed to Documentation agent
```

## Contextual References

- Session Journey template: `ejs-docs/journey/_templates/journey-template.md`
- Sub-Agent Handoff Protocol: `ejs-docs/session-lifecycle-patterns.md`
- ADR 0012: Sub-agent decision capture protocol

## Key Principle

Sub-agent decisions are first-class artifacts. Every decision a sub-agent
makes — and every handoff between sub-agents — must be attributed, recorded
with rationale, and traceable in the Session Journey.
