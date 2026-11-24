"""Command line interface: ``excel-link-finder scan`` and ``excel-link-finder clean``."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .cleaner import CLEANABLE_TYPES, clean
from .report import render_clean, render_csv, render_json, render_text
from .scanner import INFO_TYPES, ExcelLinkFinderError, ScanResult, find_workbooks, scan

EXIT_OK, EXIT_LINKS, EXIT_ERROR = 0, 1, 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-link-finder",
        description="Find and remove hidden external links in Excel workbooks (.xlsx, .xlsm, .xltx, .xltm). "
        "Works without Excel, on Windows, macOS and Linux.",
        epilog="Exit codes: 0 = no external links, 1 = external links found (or left after clean), 2 = error.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")

    scan_cmd = sub.add_parser(
        "scan",
        help="list every external link in workbooks or folders",
        description="List every place where a workbook refers to another file.",
    )
    scan_cmd.add_argument("paths", nargs="+", metavar="PATH", help="workbook files and/or folders")
    scan_cmd.add_argument("-r", "--recursive", action="store_true", help="also scan sub-folders")
    fmt = scan_cmd.add_mutually_exclusive_group()
    fmt.add_argument("--json", action="store_true", help="print JSON instead of a table")
    fmt.add_argument("--csv", action="store_true", help="print CSV, one row per finding")
    scan_cmd.add_argument("-o", "--output", metavar="FILE", help="write the report to FILE (UTF-8)")
    scan_cmd.add_argument("--no-check", action="store_true", help="do not check whether linked files exist")
    scan_cmd.add_argument("--strict", action="store_true", help="also exit with 1 for data connections and external hyperlinks")
    scan_cmd.add_argument("-q", "--quiet", action="store_true", help="print only the summary of each workbook")

    clean_cmd = sub.add_parser(
        "clean",
        help="write a copy of a workbook without its external links",
        description="Write a cleaned copy; the original file is never modified.",
    )
    clean_cmd.add_argument("input", metavar="WORKBOOK", help="workbook to clean")
    clean_cmd.add_argument("-o", "--output", metavar="FILE", help="cleaned copy (default: NAME_clean.xlsx next to the input)")
    clean_cmd.add_argument("--keep-hyperlinks", action="store_true", help="keep external hyperlinks (they are not data links)")
    clean_cmd.add_argument(
        "--only", nargs="+", metavar="TYPE", choices=CLEANABLE_TYPES, help="clean only these types: " + ", ".join(CLEANABLE_TYPES)
    )
    clean_cmd.add_argument("--force", action="store_true", help="overwrite FILE if it already exists (never the input)")
    clean_cmd.add_argument("--json", action="store_true", help="print the result as JSON")
    return parser


def _scan(args: argparse.Namespace) -> int:
    paths, skipped = find_workbooks(args.paths, args.recursive)
    results: list[ScanResult] = []
    for path in paths:
        try:
            results.append(scan(path, check_paths=not args.no_check))
        except (ExcelLinkFinderError, OSError) as exc:
            results.append(ScanResult(path, error=str(exc)))
    for path in skipped:
        print(f"skipped {path}: .xls/.xlsb files are not supported, save them as .xlsx to scan them", file=sys.stderr)
    if not results:
        print("No workbooks (.xlsx, .xlsm, .xltx, .xltm) found.", file=sys.stderr)
        return EXIT_OK
    if args.json:
        text = render_json(results, __version__)
    elif args.csv:
        text = render_csv(results)
    else:
        text = render_text(results, quiet=args.quiet)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8-sig" if args.csv else "utf-8")
        print(f"Report written to {args.output}")
    else:
        sys.stdout.write(text)
    if args.json or args.csv or args.output:
        for result in results:
            if result.error:
                print(f"error: {result.path}: {result.error}", file=sys.stderr)
    if any(r.error for r in results):
        return EXIT_ERROR
    strict = args.strict and any(f.type in INFO_TYPES for r in results for f in r.findings)
    return EXIT_LINKS if strict or any(r.has_links for r in results) else EXIT_OK


def _clean(args: argparse.Namespace) -> int:
    try:
        result = clean(args.input, args.output, keep_hyperlinks=args.keep_hyperlinks, only=args.only, overwrite=args.force)
    except (ExcelLinkFinderError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return EXIT_ERROR
    sys.stdout.write(json.dumps(result.to_dict(), indent=2) + "\n" if args.json else render_clean(result))
    return EXIT_OK if result.is_clean else EXIT_LINKS


def main(argv: list[str] | None = None) -> int:
    """Entry point of the ``excel-link-finder`` command."""
    for stream in (sys.stdout, sys.stderr):
        try:  # never crash on sheet names the console cannot display
            stream.reconfigure(errors="replace")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass
    args = build_parser().parse_args(argv)
    try:
        return _scan(args) if args.command == "scan" else _clean(args)
    except KeyboardInterrupt:
        print("interrupted", file=sys.stderr)
        return EXIT_ERROR
