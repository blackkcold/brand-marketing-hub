# v4.1 Revision Workflow

For an existing PPTX, inspect first and edit in place/copy rather than regenerate blindly.

1. Resolve the exact slide(s) and object(s) targeted by user feedback.
2. Preserve untouched slides and editable native structures.
3. Update deck_spec/coverage when content meaning or source mapping changes.
4. Re-run source/evidence checks only for affected claims/assets unless the change changes the overall decision chain.
5. Render affected slides plus neighboring slides to check visual rhythm.
6. Re-run full P0/P1 gate before final delivery.

If the runtime exposes stable slide/object IDs, persist them in revision notes. Full regeneration is reserved for explicit redesign or structural rewrite.
