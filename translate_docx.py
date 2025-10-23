#!/usr/bin/env python3
"""Translate DOCX files with Codex CLI while preserving formatting and progress."""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timedelta
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
TEXT_TAG = f"{{{W_NS}}}t"
PARA_TAG = f"{{{W_NS}}}p"
DEL_TAG = f"{{{W_NS}}}del"


@dataclass
class SegmentPart:
    node: ET.Element
    original_text: str
    preserve_space: bool


@dataclass
class Segment:
    file_name: str
    parts: List[SegmentPart]
    translate_indices: List[int]
    context: str
    segment_id: str = ""

    @property
    def char_count(self) -> int:
        return sum(len(self.parts[idx].original_text) for idx in self.translate_indices)

    def source_payload(self) -> List[Dict[str, str]]:
        return [
            {
                "id": idx,
                "text": self.parts[idx].original_text,
            }
            for idx in self.translate_indices
        ]


def leading_ws(text: str) -> str:
    idx = 0
    while idx < len(text) and text[idx].isspace():
        idx += 1
    return text[:idx]


def trailing_ws(text: str) -> str:
    idx = len(text)
    while idx > 0 and text[idx - 1].isspace():
        idx -= 1
    return text[idx:]


def restore_whitespace(original: str, translation: str) -> str:
    if not original:
        return translation
    lead = leading_ws(original)
    trail = trailing_ws(original)
    core = translation.strip()
    return f"{lead}{core}{trail}"


def register_namespaces(xml_bytes: bytes) -> None:
    with io.BytesIO(xml_bytes) as source:
        for _, (prefix, uri) in ET.iterparse(source, events=("start-ns",)):
            ET.register_namespace(prefix or "", uri)


def parse_xml(xml_bytes: bytes) -> ET.ElementTree:
    register_namespaces(xml_bytes)
    root = ET.fromstring(xml_bytes)
    return ET.ElementTree(root)


def collect_segments(tree: ET.ElementTree, file_name: str) -> List[Segment]:
    segments: List[Segment] = []
    for paragraph in tree.iterfind(f".//{PARA_TAG}"):
        deleted_nodes = {id(node) for node in paragraph.findall(f".//{DEL_TAG}//{TEXT_TAG}")}
        parts: List[SegmentPart] = []
        translate_indices: List[int] = []
        for node in paragraph.findall(f".//{TEXT_TAG}"):
            if id(node) in deleted_nodes:
                continue
            text = node.text or ""
            preserve_space = node.get(f"{{{XML_NS}}}space") == "preserve"
            part_index = len(parts)
            parts.append(SegmentPart(node=node, original_text=text, preserve_space=preserve_space))
            if text.strip():
                translate_indices.append(part_index)
        if not translate_indices:
            continue
        context = "".join(part.original_text for part in parts).strip()
        segments.append(
            Segment(
                file_name=file_name,
                parts=parts,
                translate_indices=translate_indices,
                context=context,
            )
        )
    return segments


def find_translatable_parts(extracted_dir: Path) -> Dict[str, ET.ElementTree]:
    trees: Dict[str, ET.ElementTree] = {}
    for xml_path in extracted_dir.rglob("*.xml"):
        relative = xml_path.relative_to(extracted_dir).as_posix()
        if "/_rels/" in relative:
            continue
        if not relative.startswith("word/"):
            continue
        data = xml_path.read_bytes()
        if b"<w:t" not in data:
            continue
        trees[relative] = parse_xml(data)
    return trees


def call_codex(prompt: str, codex_home: Optional[Path], timeout: int) -> str:
    if shutil.which("codex") is None:
        raise RuntimeError("codex CLI not found in PATH")
    if codex_home is not None:
        os.makedirs(codex_home, exist_ok=True)
    tmp_dir = str(codex_home) if codex_home is not None else None
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", dir=tmp_dir) as tmp:
        last_message_path = Path(tmp.name)
    cmd = [
        "codex",
        "exec",
        "--skip-git-repo-check",
        "--sandbox",
        "workspace-write",
        "--output-last-message",
        str(last_message_path),
        prompt,
    ]
    env = os.environ.copy()
    if codex_home is not None:
        env["CODEX_HOME"] = str(codex_home)
    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=timeout,
        env=env,
    )
    stdout = proc.stdout.strip()
    stderr = proc.stderr.strip()
    if proc.returncode != 0:
        raise RuntimeError(f"codex exec failed ({proc.returncode}): {stderr or stdout}")
    try:
        content = last_message_path.read_text(encoding="utf-8").strip()
    finally:
        last_message_path.unlink(missing_ok=True)
    if not content:
        raise RuntimeError(f"codex returned no content; stdout was: {stdout}")
    return content


def extract_json_block(text: str) -> str:
    stripped = text.strip()
    if stripped.startswith("```"):
        lines = stripped.splitlines()
        closing_index = None
        for idx, line in enumerate(lines[1:], start=1):
            if line.strip().startswith("```"):
                closing_index = idx
                break
        if closing_index is not None:
            stripped = "\n".join(lines[1:closing_index]).strip()
    if stripped.startswith("{") and stripped.endswith("}"):
        return stripped
    start = stripped.find("{")
    end = stripped.rfind("}")
    if start != -1 and end != -1 and end > start:
        return stripped[start : end + 1]
    raise ValueError("No JSON object found in Codex response")




def normalize_translations(raw_items: Any, expected: int) -> List[str]:
    if not isinstance(raw_items, list):
        raise ValueError("'translations' must be a list")
    normalized: List[str] = []
    for idx, item in enumerate(raw_items):
        if isinstance(item, str):
            normalized.append(item)
            continue
        if isinstance(item, dict):
            for key in ("text", "translation", "value"):
                value = item.get(key)
                if isinstance(value, str):
                    normalized.append(value)
                    break
            else:
                raise ValueError(f"Translation item {idx} missing textual field")
            continue
        raise ValueError(f"Unsupported translation type: {type(item).__name__}")
    if len(normalized) != expected:
        raise ValueError("Number of normalized translations does not match expected count")
    return normalized


def translate_segment(
    segment: Segment,
    codex_home: Optional[Path],
    timeout: int,
    source_lang: str,
    target_lang: str,
    max_retries: int,
) -> List[str]:
    payload = {
        "source_language": source_lang,
        "target_language": target_lang,
        "context": segment.context,
        "segments": segment.source_payload(),
        "instructions": [
            "Return JSON only with the schema {\"translations\": [ ... ]}.",
            "Respect the number of segments and do not merge or split them.",
            "Preserve numbers, punctuation, and placeholders exactly as in the source unless grammar requires a change.",
            "Keep leading and trailing whitespace for each segment unchanged.",
        ],
    }
    prompt = (
        "You are a professional translator specialising in Earth Sciences related Microsoft Word documents. "
        "Translate the provided segments from {source} to {target}. "
        "Output JSON only.".format(source=source_lang, target=target_lang)
        + "\nDATA:\n"
        + json.dumps(payload, ensure_ascii=False)
    )
    attempt = 0
    while attempt <= max_retries:
        attempt += 1
        try:
            response = call_codex(prompt, codex_home=codex_home, timeout=timeout)
            json_block = extract_json_block(response)
            data = json.loads(json_block)
            translations = data.get("translations")
            normalized = normalize_translations(translations, len(segment.translate_indices))
            return normalized
        except Exception as exc:  # noqa: BLE001
            if attempt > max_retries:
                raise RuntimeError(f"Failed to translate segment in {segment.file_name}: {exc}") from exc
    raise RuntimeError("Translation retries exhausted")


def apply_translations(segment: Segment, translations: Iterable[str]) -> None:
    for idx, translated in zip(segment.translate_indices, translations):
        part = segment.parts[idx]
        new_text = restore_whitespace(part.original_text, translated) if part.preserve_space else translated
        part.node.text = new_text
        part.original_text = new_text


def repackage_docx(extracted_dir: Path, output_path: Path) -> None:
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(extracted_dir.rglob("*")):
            if path.is_dir():
                continue
            arcname = path.relative_to(extracted_dir).as_posix()
            zf.write(path, arcname)


def compute_totals(segments: List[Segment]) -> int:
    return sum(segment.char_count for segment in segments)


def format_duration(seconds: float) -> str:
    seconds = max(0, int(round(seconds)))
    minutes, secs = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"



def print_progress(processed: int, total: int, current: int, total_segments: int, start_time: float, baseline_processed: int) -> None:
    percent = 100 if total == 0 else min(100, (processed / total) * 100)
    bar_width = 20
    filled = int(round(percent / 100 * bar_width))
    bar = "#" * filled + "-" * (bar_width - filled)
    elapsed = time.monotonic() - start_time
    eta_display = "--:--:--"
    finish_display = "--:--:--"
    processed_since_resume = max(processed - baseline_processed, 0)
    if processed < total and processed_since_resume > 0 and elapsed > 0:
        rate = processed_since_resume / elapsed
        if rate > 0:
            remaining = max(total - processed, 0)
            eta_seconds = remaining / rate
            eta_display = format_duration(eta_seconds)
            finish_dt = datetime.now().replace(microsecond=0) + timedelta(seconds=eta_seconds)
            finish_display = finish_dt.strftime("%H:%M:%S")
    if processed >= total and total > 0:
        eta_display = "00:00:00"
        finish_display = datetime.now().replace(microsecond=0).strftime("%H:%M:%S")
    message = f"[{bar}] {percent:6.2f}% ({current}/{total_segments} segments) ETA {eta_display} Finish {finish_display}"
    print(message, end="\r", flush=True)


def compute_file_sha256(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def load_checkpoint(checkpoint_path: Path, expected_sha: str) -> Dict[str, List[str]]:
    if not checkpoint_path.exists():
        return {}
    try:
        data = json.loads(checkpoint_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    if data.get("input_sha") != expected_sha:
        return {}
    translations = data.get("translations")
    if not isinstance(translations, dict):
        return {}
    result: Dict[str, List[str]] = {}
    for key, value in translations.items():
        if isinstance(key, str) and isinstance(value, list):
            result[key] = [str(item) for item in value]
    return result


def save_checkpoint(
    checkpoint_path: Path,
    input_sha: str,
    source_lang: str,
    target_lang: str,
    translations: Dict[str, List[str]],
) -> None:
    payload: Dict[str, Any] = {
        "input_sha": input_sha,
        "source_language": source_lang,
        "target_language": target_lang,
        "translations": translations,
    }
    checkpoint_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = checkpoint_path.with_suffix(checkpoint_path.suffix + ".tmp")
    tmp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(checkpoint_path)


def translate_docx(
    input_path: Path,
    output_path: Path,
    source_lang: str,
    target_lang: str,
    timeout: int,
    codex_home: Optional[Path],
    max_retries: int,
    dry_run: bool,
    checkpoint_path: Path,
) -> None:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    input_sha = compute_file_sha256(input_path)
    with tempfile.TemporaryDirectory() as tmp_dir_str:
        tmp_dir = Path(tmp_dir_str)
        with zipfile.ZipFile(input_path, "r") as original_zip:
            original_zip.extractall(tmp_dir)
        trees = find_translatable_parts(tmp_dir)
        if not trees:
            raise RuntimeError("No translatable XML parts found in DOCX")
        file_segments: Dict[str, List[Segment]] = {}
        for relative, tree in trees.items():
            segments = collect_segments(tree, relative)
            if segments:
                file_segments[relative] = segments
        if not file_segments:
            raise RuntimeError("No text segments requiring translation were found")
        ordered_segments: List[Segment] = []
        for relative in sorted(file_segments):
            segments = file_segments[relative]
            for index, segment in enumerate(segments):
                segment.segment_id = f"{relative}|{index}"
                ordered_segments.append(segment)
        total_chars = compute_totals(ordered_segments)
        total_segments = len(ordered_segments)
        print(
            f"Discovered {total_segments} segments across {len(file_segments)} XML parts (approx. {total_chars} characters)."
        )
        if dry_run:
            print("Dry run enabled; exiting before translation.")
            return
        translation_map = load_checkpoint(checkpoint_path, expected_sha=input_sha)
        processed_chars = 0
        completed_segments = 0
        for segment in ordered_segments:
            saved = translation_map.get(segment.segment_id)
            if not saved:
                continue
            apply_translations(segment, saved)
            processed_chars += segment.char_count
            completed_segments += 1
        baseline_processed = processed_chars
        start_time = time.monotonic()
        print_progress(processed_chars, total_chars, completed_segments, total_segments, start_time, baseline_processed)
        all_segments_done = processed_chars >= total_chars
        if not all_segments_done:
            try:
                for segment in ordered_segments:
                    if segment.segment_id in translation_map:
                        continue
                    try:
                        translations = translate_segment(
                            segment,
                            codex_home=codex_home,
                            timeout=timeout,
                            source_lang=source_lang,
                            target_lang=target_lang,
                            max_retries=max_retries,
                        )
                    except Exception as exc:  # noqa: BLE001
                        save_checkpoint(
                            checkpoint_path,
                            input_sha=input_sha,
                            source_lang=source_lang,
                            target_lang=target_lang,
                            translations=translation_map,
                        )
                        print(
                            f"\nEncountered an error on segment {segment.segment_id}: {exc}\nProgress saved to {checkpoint_path}."
                        )
                        raise
                    apply_translations(segment, translations)
                    final_texts = [segment.parts[idx].node.text or "" for idx in segment.translate_indices]
                    translation_map[segment.segment_id] = final_texts
                    processed_chars += segment.char_count
                    completed_segments += 1
                    print_progress(
                        processed_chars,
                        total_chars,
                        completed_segments,
                        total_segments,
                        start_time,
                        baseline_processed,
                    )
                    save_checkpoint(
                        checkpoint_path,
                        input_sha=input_sha,
                        source_lang=source_lang,
                        target_lang=target_lang,
                        translations=translation_map,
                    )
            except Exception:
                raise
        print()
        for relative, tree in trees.items():
            target_file = tmp_dir / relative
            tree.write(target_file, encoding="utf-8", xml_declaration=True)
        repackage_docx(tmp_dir, output_path)
        print(f"Translation complete. Output saved to {output_path}")


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Translate DOCX files with Codex CLI")
    parser.add_argument("input", type=Path, help="Path to the source DOCX file")
    parser.add_argument("--output", type=Path, help="Path for the translated DOCX")
    parser.add_argument("--source-lang", default="Catalan", help="Source language name")
    parser.add_argument("--target-lang", required=True, help="Target language name")
    parser.add_argument("--timeout", type=int, default=600, help="Codex call timeout in seconds")
    parser.add_argument("--codex-home", type=Path, help="Custom CODEX_HOME directory")
    parser.add_argument("--max-retries", type=int, default=2, help="Retries per segment on Codex errors")
    parser.add_argument("--dry-run", action="store_true", help="List segments without translating")
    parser.add_argument(
        "--checkpoint",
        type=Path,
        help="Path for persisting translation progress (defaults next to the input file)",
    )
    return parser.parse_args(argv)


def main(argv: List[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    input_path: Path = args.input
    if args.output:
        output_path = args.output
    else:
        output_path = input_path.with_name(f"{input_path.stem}_translated{input_path.suffix}")
    if args.checkpoint:
        checkpoint_path = args.checkpoint
    else:
        checkpoint_path = input_path.with_suffix(f"{input_path.suffix}.progress.json")
    if args.codex_home is not None:
        codex_home: Optional[Path] = args.codex_home
    else:
        env_home = os.environ.get("CODEX_HOME")
        codex_home = Path(env_home) if env_home else None
    try:
        translate_docx(
            input_path=input_path,
            output_path=output_path,
            source_lang=args.source_lang,
            target_lang=args.target_lang,
            timeout=args.timeout,
            codex_home=codex_home,
            max_retries=args.max_retries,
            dry_run=args.dry_run,
            checkpoint_path=checkpoint_path,
        )
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
