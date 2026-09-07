# dsh-law-thesis-review

A DeepSeek Harness **skill plugin** that reviews student law theses (国际经济法/法学论文) and outputs a Word review draft with **tracked changes + Word comments**.

## What it does

- Reads a student Word (.docx) thesis (single file or a folder).
- Checks against a bundled common-issue checklist (硕士/本科自查清单合并版 + 国际经济法/法学论文高发问题).
- Applies definite fixes as **Word tracked changes** (w:ins / w:del).
- Writes **Word comments** for issues the student must handle themselves (overall section problems, missing basis/clarification, etc.).
- Emits a short, objective overall review + a fix list.

## Install (local / from npm)

~~~bash
# local path
dsh plugin --profile web add link:D:/Coding/Creating/dsh-law-thesis-review
# or from npm (once published)
dsh plugin --profile web add @signalight/dsh-law-thesis-review
~~~

Restart the web profile. The skill is then auto-discovered; trigger it with phrases like “批改学生论文”, “审阅这篇论文”, “以修订+批注检查论文”.

## Usage

The skill runs the bundled script:

~~~bash
# 1) extract text + paragraph indices
python skills/iet-law-thesis-review/scripts/review_docx.py extract 论文.docx --out ./work

# 2) review: apply tracked edits + comments from a spec
python skills/iet-law-thesis-review/scripts/review_docx.py review 论文.docx --spec spec.json --out 论文_审查稿.docx [--author "教师名"]
~~~

Spec shape shown in `skills/iet-law-thesis-review/examples/spec-template.json`.

## Notes

- **Self-contained**: `apply_edits.py` / `models.py` / `add_comment.py` are vendored; the skill does **not** require an external word-docx plugin.
- Checklist is **bundled** (`skills/iet-law-thesis-review/references/checklist.md`); no dependency on any local Obsidian vault path.
- The output preserves the original title and any student info; revision/comment author defaults to the local Office/WPS username or can be set via `--author`.
- Word: open the `_审查稿.docx` to see the redline + comment bubbles.

## License

MIT

