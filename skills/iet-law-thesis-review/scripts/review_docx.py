#!/usr/bin/env python
"""iet-law-thesis-review automation.

Subcommands:
  extract  INPUT.docx --out OUTDIR
             -> paragraphs.json / document.md / footnotes.json / meta.json
  review   INPUT.docx --spec SPEC.json --out OUTPUT.docx [--author AUTHOR]

SPEC.json shape:
{
  "summary": "简短、客观、以指出问题为主的总体评语（可选）",
  "edits":   [ {"operation":"replace","paragraph_index":165,"old_text":"利用","new_text":"运用"}, ... ],
  "comments":[ {"paragraph_index":83,"anchor_text":"美国","text":"请核实这里的表述。"}, ... ]
}
Edits are applied as Word tracked changes (w:ins / w:del). Comments are inserted
as real Word comments. The output keeps the original title and all student info.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

# The paragraph-index semantics match apply_edits: direct children of w:body only.
# apply_edits / models / add_comment are vendored in this same scripts/ directory,
# so the skill is self-contained (no external word-docx plugin required).
# The paragraph-index semantics match apply_edits: direct children of w:body only.
from apply_edits import W_NS, apply_edits as _apply_tracked_edits  # noqa: E402
from models import EditOperation  # noqa: E402
from add_comment import _insert_new_comment_markers  # noqa: E402

XML_NS = "http://www.w3.org/XML/1998/namespace"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
COMMENTS_REL_TYPE = DOC_REL_NS + "/comments"
COMMENTS_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)


def _t(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _para_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(_t("t")))


def _read_zip(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {n: zf.read(n) for n in zf.namelist()}


def _write_zip(path: Path, entries: dict[str, bytes]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            zf.writestr(name, data)


def _detect_author(input_path: Path) -> str:
    """Try, in order: registry username -> input docx lastModifiedBy != creator -> ''"""
    # 1) Windows registry (Office / WPS)
    try:
        import winreg
        registry_paths = [
            r"Software\Microsoft\Office\16.0\Common\UserInfo",
            r"Software\Microsoft\Office\16.0\Word\Options",
            r"Software\Microsoft\Office\15.0\Common\UserInfo",
            r"Software\Microsoft\Office\15.0\Word\Options",
            r"Software\Kingsoft\Office\6.0\common\UserInfo",
            r"Software\Kingsoft\Office\6.0\user",
        ]
        for rp in registry_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rp) as k:
                    try:
                        val, _ = winreg.QueryValueEx(k, "UserName")
                    except OSError:
                        val = None
                    if not val:
                        try:
                            val, _ = winreg.QueryValueEx(k, "UserInitials")
                        except OSError:
                            val = None
                    if val:
                        return str(val).strip()
            except OSError:
                continue
    except Exception:
        pass

    # 2) input docx: prefer lastModifiedBy if it differs from the original creator
    try:
        entries = _read_zip(input_path)
        if "docProps/core.xml" in entries:
            root = etree.fromstring(entries["docProps/core.xml"])
            ns = {"dc": "http://purl.org/dc/elements/1.1/",
                  "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"}
            creator = (root.findtext("dc:creator", namespaces=ns) or "").strip()
            last = (root.findtext("cp:lastModifiedBy", namespaces=ns) or "").strip()
            if last and last != creator:
                return last
            if creator:
                return creator
    except Exception:
        pass

    # 3) empty (leave blank)
    return ""


def _resolve_author(input_path: Path, explicit: str | None) -> str:
    if explicit is not None:
        return explicit
    return _detect_author(input_path)


# --- extraction ---------------------------------------------------------------
def cmd_extract(input_path: Path, outdir: Path) -> int:
    entries = _read_zip(input_path)
    doc = etree.fromstring(entries["word/document.xml"])
    body = doc.find(f".//{_t('body')}")
    if body is None:
        print("ERROR: no w:body found.", file=sys.stderr)
        return 1
    paras = body.findall(_t("p"))
    records = []
    for i, p in enumerate(paras):
        records.append({"index": i, "text": _para_text(p)})

    outdir.mkdir(parents=True, exist_ok=True)
    (outdir / "paragraphs.json").write_text(
        json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    lines = []
    for rec in records:
        lines.append(f"[P{rec['index']:04d}] {rec['text']}")
    (outdir / "document.md").write_text(
        "\n\n".join(lines), encoding="utf-8"
    )

    # footnotes
    footnote_list = []
    if "word/footnotes.xml" in entries:
        fn = etree.fromstring(entries["word/footnotes.xml"])
        for f in fn.iter(_t("footnote")):
            fid = f.get(_t("id"))
            text = _para_text(f)
            if text.strip():
                footnote_list.append({"id": fid, "text": text})
    (outdir / "footnotes.json").write_text(
        json.dumps(footnote_list, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # meta
    meta = {
        "input_file": str(input_path),
        "paragraph_count": len(paras),
        "footnote_count": len(footnote_list),
    }
    if "docProps/core.xml" in entries:
        root = etree.fromstring(entries["docProps/core.xml"])
        ns = {"dc": "http://purl.org/dc/elements/1.1/",
              "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"}
        meta["creator"] = root.findtext("dc:creator", namespaces=ns)
        meta["lastModifiedBy"] = root.findtext("cp:lastModifiedBy", namespaces=ns)
        meta["title"] = root.findtext("dc:title", namespaces=ns) or ""
    (outdir / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"Extracted {len(paras)} paragraphs, {len(footnote_list)} footnotes -> {outdir}")
    return 0


# --- tracked edits ------------------------------------------------------------
def _apply_edits(input_path: Path, edits: list[dict], intermediate: Path,
                 author: str) -> list[str]:
    ops = []
    for e in edits:
        ops.append(EditOperation(
            operation=e["operation"],
            paragraph_index=int(e["paragraph_index"]),
            old_text=e.get("old_text"),
            new_text=e.get("new_text"),
            author=author,
        ))
    if not ops:
        shutil.copyfile(input_path, intermediate)
        return []
    diags = _apply_tracked_edits(input_path, ops, intermediate)
    return [d.message for d in diags if d.level in ("warning", "error", "info")]


# --- comments ----------------------------------------------------------------
def _ensure_comment_infra(entries: dict[str, bytes]) -> None:
    ct_name = "[Content_Types].xml"
    ct = etree.fromstring(entries[ct_name])
    if not any(o.get("PartName") == "/word/comments.xml"
               for o in ct.iter(f"{{{CT_NS}}}Override")):
        o = etree.SubElement(ct, f"{{{CT_NS}}}Override")
        o.set("PartName", "/word/comments.xml")
        o.set("ContentType", COMMENTS_CT)
    entries[ct_name] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    rels_name = "word/_rels/document.xml.rels"
    rels = etree.fromstring(entries[rels_name])
    if not any(r.get("Type") == COMMENTS_REL_TYPE
               for r in rels.iter(f"{{{REL_NS}}}Relationship")):
        maxrid = 0
        for r in rels.iter(f"{{{REL_NS}}}Relationship"):
            rid = r.get("Id")
            if rid and rid.startswith("rId"):
                try:
                    maxrid = max(maxrid, int(rid[3:]))
                except ValueError:
                    pass
        rel = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
        rel.set("Id", f"rId{maxrid + 1}")
        rel.set("Type", COMMENTS_REL_TYPE)
        rel.set("Target", "comments.xml")
    entries[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    if "word/comments.xml" not in entries:
        root = etree.Element(_t("comments"))
        entries["word/comments.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _add_comments(input_path: Path, comments: list[dict], output_path: Path,
                  author: str) -> tuple[int, list[str]]:
    """Insert Word comments (range markers + comments.xml). Returns (added, skips)."""
    entries = _read_zip(input_path)
    doc = etree.fromstring(entries["word/document.xml"])
    body = doc.find(f".//{_t('body')}")
    par_count = len(body.findall(_t("p")))

    if "word/comments.xml" in entries:
        comments_root = etree.fromstring(entries["word/comments.xml"])
        max_id = -1
        for c in comments_root.iter(_t("comment")):
            wid = c.get(_t("id"))
            if wid is not None:
                try:
                    max_id = max(max_id, int(wid))
                except ValueError:
                    pass
    else:
        comments_root = etree.Element(_t("comments"))
        max_id = -1

    _ensure_comment_infra(entries)

    added = 0
    skips: list[str] = []
    next_id = max_id + 1
    for c in comments:
        para_idx = int(c["paragraph_index"])
        anchor = c.get("anchor_text") or ""
        text = c.get("text") or ""
        if para_idx < 0 or para_idx >= par_count:
            skips.append(f"comment: paragraph_index {para_idx} out of range (doc has {par_count})")
            continue
        new_id = str(next_id)
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=False) as tf:
            tmp_path = Path(tf.name)
            tf.write(etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True))
        ok = False
        try:
            ok = _insert_new_comment_markers(tmp_path, para_idx, anchor, new_id)
            if ok:
                doc = etree.fromstring(tmp_path.read_bytes())
        except Exception:
            ok = False
        finally:
            tmp_path.unlink(missing_ok=True)
        if not ok:
            skips.append(f"comment id {new_id}: anchor text {anchor!r} not found in paragraph {para_idx}")
            continue

        c_el = etree.SubElement(comments_root, _t("comment"))
        c_el.set(_t("id"), new_id)
        c_el.set(_t("author"), author or "")
        c_el.set(_t("date"), datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))
        if author:
            initials = "".join(ch[0] for ch in author.split()) or "T"
            c_el.set(_t("initials"), initials[:4])
        p = etree.SubElement(c_el, _t("p"))
        r = etree.SubElement(p, _t("r"))
        t = etree.SubElement(r, _t("t"))
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text
        added += 1
        next_id += 1

    entries["word/document.xml"] = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    entries["word/comments.xml"] = etree.tostring(comments_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _write_zip(output_path, entries)
    return added, skips


# --- summary markdown ---------------------------------------------------------
def _summary_md(input_path: Path, spec: dict, out_md: Path, author: str,
                added_comments: int, skipped: list[str], skipped_edits: list[str] | None = None) -> None:
    name = input_path.name
    lines = [f"# 审查说明：{name}", ""]
    if spec.get("summary"):
        lines.append("## 总体评语")
        lines.append(spec["summary"])
        lines.append("")
    if spec.get("edits"):
        lines.append("## 待修改（已以修订模式写入）")
        lines.append("")
        for e in spec["edits"]:
            op = e.get("operation", "")
            pi = e.get("paragraph_index")
            old = e.get("old_text")
            new = e.get("new_text")
            if op == "replace":
                lines.append(f"- [P{pi}] {old} → {new}")
            elif op == "delete":
                lines.append(f"- [P{pi}] 删除：{old}")
            elif op == "insert":
                lines.append(f"- [P{pi}] 插入：{new}")
        lines.append("")
    if spec.get("comments"):
        lines.append("## 批注（已以Word批注形式写入）")
        lines.append("")
        for c in spec["comments"]:
            lines.append(f"- [P{c.get('paragraph_index')}] {c.get('text')}（锚定：{c.get('anchor_text')}）")
        lines.append("")
    if skipped_edits:
        lines.append("## 未能写入的修订")
        lines.append("")
        for s in skipped_edits:
            lines.append(f"- {s}")
        lines.append("")
    if skipped:
        lines.append("## 未能写入的批注")
        lines.append("")
        for s in skipped:
            lines.append(f"- {s}")
        lines.append("")
    lines.append(f"批注/修订作者：{author or '（留空）'}")
    lines.append("")
    out_md.write_text("\n".join(lines), encoding="utf-8")


# --- review command -----------------------------------------------------------
def cmd_review(input_path: Path, spec_path: Path, out_path: Path, author_override: str | None) -> int:
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    author = _resolve_author(input_path, spec.get("author") or author_override)

    edits = spec.get("edits", [])
    comments = spec.get("comments", [])

    with tempfile.TemporaryDirectory(dir=out_path.parent) as tmp:
        intermediate = Path(tmp) / "_intermediate.docx"
        edit_messages = _apply_edits(input_path, edits, intermediate, author)
        if os.environ.get("THESIS_REVIEW_DEBUG"):
            (out_path.parent / "edit_messages.debug.txt").write_text(
                "\n".join(edit_messages), encoding="utf-8")
        added_comments, skips = _add_comments(intermediate, comments, out_path, author)

    applied_summary = next((m for m in edit_messages if "Applied" in m), "")
    failed_edits = [m for m in edit_messages
                    if ("not found" in m or "out of range" in m or m.startswith("ERROR"))]

    md_out = out_path.with_suffix(".审查说明.md")
    _summary_md(input_path, spec, md_out, author, added_comments, skips, failed_edits)

    print(f"Output: {out_path}")
    print(f"Summary: {md_out}")
    print(f"Tracked edits: {applied_summary}")
    print(f"Comments added: {added_comments}")
    if failed_edits:
        print("Failed edits (skipped):")
        for s in failed_edits:
            print(f"  - {s}")
    if skips:
        print("Skipped comments:")
        for s in skips:
            print(f"  - {s}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="iet-law-thesis-review automation")
    sub = parser.add_subparsers(dest="command", required=True)

    pe = sub.add_parser("extract", help="Extract structured text + paragraph indices")
    pe.add_argument("input_docx", type=Path)
    pe.add_argument("--out", type=Path, required=True)

    pr = sub.add_parser("review", help="Apply tracked edits + comments to a docx")
    pr.add_argument("input_docx", type=Path)
    pr.add_argument("--spec", type=Path, required=True)
    pr.add_argument("--out", type=Path, required=True)
    pr.add_argument("--author", default=None)

    args = parser.parse_args(argv)
    if args.command == "extract":
        return cmd_extract(args.input_docx, args.out)
    if args.command == "review":
        return cmd_review(args.input_docx, args.spec, args.out, args.author)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
