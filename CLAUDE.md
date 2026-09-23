# Claude Instructions — Access2SQL

## Development log
After every session (or at the end of a conversation), append a summary of what was
discussed, decided, and changed to:

    src/claude_dev_log.md

Each entry must include:
- **Problem / request** — what the user asked for
- **Investigation** — how the issue was diagnosed or the approach was chosen
- **Reasoning** — why this solution over alternatives
- **Changes made** — files edited/created and what changed in them
- **Version bump** — if VERSION was changed, note old → new

Use the same heading/section style already in the file (H2 for session, H3 for subsections,
tables and fenced code blocks for detail).

## Versioning
The script uses a single `VERSION` constant in `access2sql.py`.
- Patch fix (bug, typo, minor tweak) → increment last digit, e.g. `1.1` → `1.1.1`
- New feature → increment middle digit, e.g. `1.1` → `1.2`
- Breaking change → increment major digit, e.g. `1.1` → `2.0`

Always update the version when making any code change.
