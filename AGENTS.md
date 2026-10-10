# TNP agent instructions

1. Read `.specify/memory/constitution.md`, then `docs/SPECKIT_ARENA.md` before work.
2. Use the Spec Kit feature workflow for new bounded features and its bug workflow for
   defects. If native invocation is unavailable, read and follow the checked-in skill
   instructions and disclose fallback execution as documented in the guide.
3. Do not implement before specification/assessment and planning unless the user explicitly
   requests an emergency minimal change. Record that exception and missing evidence.
4. Existing tests and TNP compatibility behavior are authoritative. Do not weaken tests to
   get green output. Do not backfill historical PROMPT-001 through PROMPT-030 specs.
5. Linux results are development evidence, never Windows, PowerPoint COM, or Portable
   acceptance. Report unperformed checks honestly.
6. Confirm the authoritative base; work only on the assigned Arena session branch. Commit
   scoped artifacts and authorized changes, push that branch to Git, and open the PR
   against the requested base. Stage explicit paths, never `git add .`.
7. Do not automatically merge acceptance PRs, publish, or bump application version/build.
   Keep unrelated PRs untouched and preserve the immutable stable backup.
