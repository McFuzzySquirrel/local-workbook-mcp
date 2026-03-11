session_id:
author:
date:
repo:
branch:
agents_involved:
decision_detected: false
adr_links: []
tags: []
refs: []

> Save as: `ejs-docs/journey/YYYY/ejs-session-YYYY-MM-DD-<seq>.md`

# Problem / Intent
Describe the purpose of the session.

# Interaction Summary (Required)
Capture the human↔agent collaboration trail in a concise, chronological form.

Recommended format:
- Human: <prompt / request>
  - Agent [agent-name]: <response / guidance>
  - Outcome: <what changed / what was decided>

Do:
- Attribute every entry by agent name (e.g., `Agent [copilot]`, `Agent [claude]`, `Agent [ejs-journey]`).
- Include key prompts, constraints, and "why" questions.
- Include pivotal corrections/pivots and what triggered them.
- Record sub-agent delegations inline (e.g., "Agent [copilot] delegated testing to sub-agent [test-agent]").

Avoid:
- Pasting full transcripts.
- Including secrets, tokens, or sensitive content.
- Omitting agent names — every action must be attributed.

# Agent Collaboration Summary
Which agents participated (by name):
Key suggestions made (and by which agent):
Corrections applied by human:

# Sub-Agent Contributions
When any agent delegates work to sub-agents, capture each sub-agent's contribution here (attributed by name).

## Sub-Agent: <agent-name / agent-type>
- **Task delegated:** What was the sub-agent asked to do?
- **Decisions made:** What did the sub-agent decide (e.g., approach chosen, trade-offs weighed)?
- **Alternatives considered:** What other approaches did the sub-agent evaluate?
- **Outcome:** What was the result?
- **Handoff to other agents:** Did this sub-agent's output feed into another sub-agent's work? If so, describe the handoff.

(Repeat for each sub-agent involved. Remove this section if no sub-agents were used.)

# Agent Influence (Required)
Track how each agent influenced the outcome (attribute by name).
- Suggestions adopted (which agent, what suggestion):
- Suggestions rejected (which agent, what suggestion, why rejected):
- Human overrides / corrections (what was changed, which agent was overridden):

# Experiments / Evidence
What was tried?  
What happened?  
What evidence changed your mind?  

# Iteration Log
Notable pivots, reversals, or refinements (and why).

# Decisions Made
- Decision:
  - Reason:
  - Impact:

# Key Learnings
Technical insights  
Prompting insights  
Tooling insights  

# If Repeating This Work
Do this:  
Avoid this:  
Watch out for:  

# Future Agent Guidance
Prefer:  
Avoid:  

## MACHINE EXTRACTS

### INTERACTION_EXTRACT

### DECISIONS_EXTRACT

### LEARNING_EXTRACT

### AGENT_GUIDANCE_EXTRACT

### SUB_AGENT_EXTRACT
