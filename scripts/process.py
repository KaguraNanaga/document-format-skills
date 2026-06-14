#!/usr/bin/env python3
"""One-shot CLI pipeline for document formatting tasks."""

import argparse
import json
import logging
import os
import shutil
import sys
import tempfile
from copy import deepcopy
from pathlib import Path

from docx import Document

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.analyzer import (  # noqa: E402
    analyze_font,
    analyze_numbering,
    analyze_paragraph_format,
    analyze_punctuation,
    print_report,
)
from scripts.formatter import (  # noqa: E402
    PRESETS,
    _merge_preset_settings,
    format_document,
    load_custom_preset,
    load_preset_file,
)
from scripts.punctuation import process_document as fix_punctuation  # noqa: E402


logger = logging.getLogger("docformat.process")


def analyze_document(path, as_json=False):
    doc = Document(path)
    results = {
        "punctuation": analyze_punctuation(doc),
        "numbering": analyze_numbering(doc),
        "paragraph": analyze_paragraph_format(doc),
        "font": analyze_font(doc),
    }
    if as_json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        print_report(results)
    return results


def _format_overrides(args):
    overrides = {}
    if args.deep_clean:
        overrides["deep_clean"] = True
    if args.split_heading_at_punct:
        overrides["split_heading_at_punct"] = True
    if args.smart_table_align:
        overrides.setdefault("table", {})["smart_align"] = True
    if args.no_page_number:
        overrides["page_number"] = False
    if args.no_replace_existing_page_number:
        overrides["replace_existing_page_number"] = False

    page_number_fields = {
        "page_number_style": args.page_number_style,
        "page_number_position": args.page_number_position,
        "page_number_offset_mm": args.page_number_offset_mm,
        "page_number_size": args.page_number_size,
        "page_number_font": args.page_number_font,
    }
    for key, value in page_number_fields.items():
        if value is not None:
            overrides[key] = value

    return overrides


def _load_cli_custom_settings(args):
    settings = load_preset_file(args.custom_settings) if args.custom_settings else None
    overrides = _format_overrides(args)
    if args.preset == "custom" and overrides and settings is None:
        settings = load_custom_preset() or deepcopy(PRESETS["official"])
    if overrides:
        settings = _merge_preset_settings(settings or {}, overrides)
    return settings


def _convert_input_if_needed(path):
    ext = Path(path).suffix.lower()
    if ext not in (".doc", ".wps"):
        return path, None

    if os.name != "nt":
        raise RuntimeError(".doc/.wps conversion is only available on Windows with Word or WPS installed")

    from scripts.converter import convert_to_docx

    logger.info("Converting %s to .docx...", ext)
    temp_docx = convert_to_docx(path)
    return temp_docx, temp_docx


def _prepare_output_path(output_path):
    ext = Path(output_path).suffix.lower()
    if ext not in (".doc", ".wps"):
        return output_path, None

    fd, temp_output = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    return temp_output, temp_output


def _convert_output_if_needed(temp_output_docx, requested_output):
    requested_ext = Path(requested_output).suffix.lower()
    if requested_ext not in (".doc", ".wps"):
        return requested_output

    if os.name != "nt":
        fallback = str(Path(requested_output).with_suffix(".docx"))
        shutil.copy2(temp_output_docx, fallback)
        logger.warning("Cannot convert to %s on this platform; saved %s", requested_ext, fallback)
        return fallback

    from scripts.converter import convert_from_docx

    try:
        return convert_from_docx(
            temp_output_docx,
            requested_output,
            format=requested_ext.lstrip("."),
        )
    except Exception as exc:
        fallback = str(Path(requested_output).with_suffix(".docx"))
        shutil.copy2(temp_output_docx, fallback)
        logger.warning("Output conversion failed: %s; saved %s", exc, fallback)
        return fallback


def run_pipeline(args):
    input_path, temp_input = _convert_input_if_needed(args.input)
    output_docx = None
    temp_output = None
    try:
        if args.mode == "analyze":
            return analyze_document(input_path, as_json=args.json)

        output_docx, temp_output = _prepare_output_path(args.output)

        if args.mode == "punctuation":
            fix_punctuation(input_path, output_docx, space_mode=args.space_mode)
        elif args.mode == "format":
            format_document(
                input_path,
                output_docx,
                preset_name=args.preset,
                revision_mode=args.revision,
                bold_serial=not args.no_bold_serial,
                custom_settings=_load_cli_custom_settings(args),
            )
        elif args.mode == "smart":
            fd, temp_fixed = tempfile.mkstemp(suffix=".docx")
            os.close(fd)
            try:
                fix_punctuation(input_path, temp_fixed, space_mode=args.space_mode)
                format_document(
                    temp_fixed,
                    output_docx,
                    preset_name=args.preset,
                    revision_mode=args.revision,
                    bold_serial=not args.no_bold_serial,
                    custom_settings=_load_cli_custom_settings(args),
                )
            finally:
                if os.path.exists(temp_fixed):
                    os.unlink(temp_fixed)
        else:
            raise ValueError(f"Unknown mode: {args.mode}")

        actual_output = _convert_output_if_needed(output_docx, args.output)
        logger.info("Saved: %s", actual_output)
        return actual_output
    finally:
        for path in (temp_input, temp_output):
            if path and os.path.exists(path):
                try:
                    os.unlink(path)
                except OSError:
                    pass


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Analyze, fix punctuation, format, or smart-process document files."
    )
    parser.add_argument("mode", choices=("smart", "analyze", "punctuation", "format"))
    parser.add_argument("input", help="Input .docx file; .doc/.wps are supported on Windows with Office/WPS")
    parser.add_argument("output", nargs="?", help="Output file for non-analyze modes")
    parser.add_argument("--json", action="store_true", help="Print analysis result as JSON")
    parser.add_argument(
        "--space-mode",
        choices=("remove_all", "keep_en_boundary", "keep_all"),
        default="remove_all",
        help="Spacing mode used during punctuation/smart processing.",
    )
    parser.add_argument(
        "--preset",
        choices=tuple(PRESETS.keys()) + ("custom",),
        default="official",
        help="Formatting preset.",
    )
    parser.add_argument("--custom-settings", help="Custom preset/config JSON file")
    parser.add_argument("--revision", action="store_true", help="Output supported formatting changes as revisions")
    parser.add_argument("--no-bold-serial", action="store_true", help="Do not bold 一是/二是 body prefixes")
    parser.add_argument("--deep-clean", action="store_true", help="Clear inherited formatting before applying preset")
    parser.add_argument("--split-heading-at-punct", action="store_true", help="Split heading+body paragraphs at punctuation")
    parser.add_argument("--smart-table-align", action="store_true", help="Align table cells by content type")
    parser.add_argument("--no-page-number", action="store_true", help="Skip adding page numbers")
    parser.add_argument(
        "--page-number-style",
        choices=("dash", "plain", "page_text", "page_total"),
        help="Page number style.",
    )
    parser.add_argument(
        "--page-number-position",
        choices=("outside", "left", "center", "right"),
        help="Page number position.",
    )
    parser.add_argument("--page-number-offset-mm", type=float, help="Page number offset below text area")
    parser.add_argument("--page-number-size", type=int, help="Page number font size")
    parser.add_argument("--page-number-font", help="Page number font")
    parser.add_argument("--no-replace-existing-page-number", action="store_true", help="Keep existing page numbers")
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        help="Logging level.",
    )
    args = parser.parse_args(argv)

    if args.mode != "analyze" and not args.output:
        parser.error("output is required for smart, punctuation, and format modes")

    logging.basicConfig(level=getattr(logging, args.log_level), format="%(message)s")
    run_pipeline(args)


if __name__ == "__main__":
    main()
