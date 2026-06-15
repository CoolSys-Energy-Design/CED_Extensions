## Profile

This is a bit about Reed Pinterich as it pertains to his working career.
Reed is a 28 almost 29 year old man born July 1st 1997, he is an ENTJ, he has worked for a plastic company as the site electrical engineer when he was 23 for about a year, he has also worked in substation design for about a year when he was 25. Now he is an AI prompt engineer working closely with Claude code to automate aspects of the MEP industry.

In his spare time Reed enjoys the hobbies of chess, he is a guitarist, and he enjoys pokemon card collecting and just a few video games.
Reed would like to Buy a Duplex next year and live in one side and rent out the other.

## Data Safety

- Always create a backup copy of any data/config file (YAML, JSON) before overwriting or transforming it. Never destroy original values without a recoverable copy.

## Mission & Scope

- This tooling is built for Reed's firm (**CoolSys**) — internal production use, not a third-party product.
- Current focus is the **electrical / power** discipline (panels, circuiting, power plans, Revit MCP automation).
- **Over time** the goal is to expand into the other MEP disciplines (mechanical, plumbing) as well — keep architecture and conventions general enough to grow into them.

## How to Work With Reed (working preferences & standards)

- **Collaborative, not autonomous.** Propose the approach and confirm major moves before executing. Reed wants to *understand the reasoning and problem-solving process* behind each step — so think out loud, explain *why*, and don't just deliver a finished result without showing the path.
- Surface trade-offs and your reasoning as you go; treat each task as a chance to make the process legible, not just to get to the answer.
- **Confirmation granularity: per-room / per-phase.** Check in at room or phase boundaries — do NOT ask before every individual family placement (that's far too granular).
- _Naming conventions and QA/QAQC standards: to be captured over time as they surface during real work._

## Workflow / Before Editing

- Before editing, confirm the exact file/version the user is actively running (e.g., the 1.0 tool vs the 2.0 panel). Verify the target file matches the running tool before making changes.

## Conventions

- Write general, reusable rules by default rather than project- or site-specific ones (e.g., do not hardcode 'Ellsworth' rules). Ask before scoping logic narrowly.

## Revit / Placement Logic

- When validating placement against a room, use the actual room polygon/boundary, not the bounding box, since rooms are frequently irregular shapes.

## Output Style

- Keep responses concise and avoid dumping large code/output blocks in a single message; chunk long outputs to stay under token limits.