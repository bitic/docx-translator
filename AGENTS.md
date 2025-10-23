# Repository Guidelines

## Project Structure & Module Organization
Keep the repository intentionally small. `translate_docx.py` contains the complete DOCX-to-DOCX translation pipeline (unzips, talks to the Codex CLI, rewrites XML, repackages). Work with your original DOCX files and, if helpful, keep a lightweight sample document for quick verification. Directories such as `__pycache__/` are derived and can be regenerated. The Codex CLI relies on the credentials stored in `~/.codex`; create a dedicated directory and use `--codex-home` if you need a separate configuration.

## Build, Test, and Development Commands
This project runs on Python >= 3.11 and only uses the standard library. Useful commands:

```bash
python translate_docx.py sample.docx --target-lang "Spanish" --dry-run
python translate_docx.py source.docx --target-lang "English"
python translate_docx.py source.docx --target-lang "English" --checkpoint progress.json
python -m compileall translate_docx.py
```

The first command inventories segments without consuming quota. The second performs a full translation while reusing the current `~/.codex`. The third specifies a custom progress file when you need to resume across sessions. The last confirms that the code compiles before committing.

## Coding Style & Naming Conventions
Follow PEP 8: four-space indentation, `snake_case` for functions and variables, `UpperCamelCase` for dataclasses. Prefer small helpers with descriptive names (`normalize_translations`, `print_progress`). Keep code in ASCII unless DOCX interoperability requires UTF-8. Add comments only when the XML flow or Codex interaction is not obvious.

## Testing Guidelines
Before substantial changes, run the `--dry-run` mode and check the segment and character counts. To validate XML transformations, temporarily extract the DOCX (`unzip -d tmp/`) and inspect the affected `w:t` nodes. When introducing new Codex logic, test with a small sample fixture to validate checkpointing, the progress bar with ETA/end time, and DOCX writing. Plan automated tests (CLI mocks) in `tests/` once the interface stabilizes.

## Commit & Pull Request Guidelines
Write short present-tense commit messages, e.g., `feat: add ETA to progress bar`. Each PR should explain what changed, how to reproduce it (`python translate_docx.py ...`), and any quota or credential implications. Attach artifacts (DOCX/PDF) when formatting may vary. Keep functionality and documentation changes separate to simplify review.

## Operational Notes
CLI calls run segment by segment and may hit quota or credential limits; the automatic checkpoint is stored as `*.progress.json` and is no longer deleted after translation (inspect or remove it whenever appropriate). If Codex requires reauthentication, run `codex login` and re-run the command: the script resumes from the last checkpoint. The progress bar prints the ETA and estimated completion time based on processed characters; leave enough buffer height so the final line renders cleanly.
