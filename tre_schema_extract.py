#!/usr/bin/env python3
"""Privacy-conscious, offline schema extractor for CSV/XLSX files."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import fnmatch
import hashlib
import json
import math
import re
import sys
from collections import Counter
from pathlib import Path

VERSION = "1.0.0"
ISO_DATE = re.compile(r"^\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?$")


def load_config(path):
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    if cfg.get("privacy_mode", "schema_only") not in {"schema_only", "aggregate_profile"}:
        raise ValueError("privacy_mode must be schema_only or aggregate_profile")
    return cfg


def matches_any(path, patterns):
    p = path.as_posix()
    return any(fnmatch.fnmatch(p, pattern) for pattern in patterns)


def logical_table(relative_path, cfg):
    p = relative_path.as_posix()
    for rule in cfg.get("table_rules", []):
        if fnmatch.fnmatch(p, rule["path_glob"]):
            return rule["name"]
    return str(relative_path.with_suffix("")).replace("\\", "/")


def infer_scalar(value):
    v = value.strip()
    low = v.lower()
    if low in {"true", "false", "yes", "no"}:
        return "boolean"
    try:
        int(v)
        return "integer"
    except ValueError:
        pass
    try:
        x = float(v)
        if math.isfinite(x):
            return "float"
    except ValueError:
        pass
    if ISO_DATE.match(v):
        try:
            dt.datetime.fromisoformat(v.replace("Z", "+00:00"))
            return "datetime" if ("T" in v or " " in v) else "date"
        except ValueError:
            pass
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            dt.datetime.strptime(v, fmt)
            return "date"
        except ValueError:
            pass
    return "string"


def resolved_type(counter):
    kinds = set(counter)
    if not kinds:
        return "unknown"
    if kinds <= {"integer"}:
        return "integer"
    if kinds <= {"integer", "float"}:
        return "float"
    if kinds <= {"date", "datetime"}:
        return "datetime" if "datetime" in kinds else "date"
    if kinds <= {"boolean"}:
        return "boolean"
    return "string"


def open_csv(path, cfg):
    last = None
    for enc in cfg.get("csv_encoding_candidates", ["utf-8-sig", "utf-8", "cp1252"]):
        try:
            f = open(path, encoding=enc, newline="")
            sample = f.read(65536)
            f.seek(0)
            delimiter = cfg.get("delimiter", "auto")
            if delimiter == "auto":
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
                    reader = csv.reader(f, dialect=dialect)
                except csv.Error:
                    # Sniffer commonly fails on empty/single-column data
                    # dictionaries. Prefer a delimiter seen consistently in
                    # the first non-empty lines; otherwise treat each line as
                    # one field instead of skipping the file.
                    lines = [line for line in sample.splitlines() if line.strip()][:20]
                    candidates = [",", "\t", ";", "|"]
                    scores = {
                        candidate: sum(line.count(candidate) for line in lines)
                        for candidate in candidates
                    }
                    fallback = max(candidates, key=lambda candidate: scores[candidate])
                    if scores[fallback] == 0:
                        fallback = ","
                    reader = csv.reader(f, delimiter=fallback)
                return f, reader, enc
            return f, csv.reader(f, delimiter=delimiter), enc
        except UnicodeDecodeError as e:
            last = e
            try:
                f.close()
            except Exception:
                pass
    raise RuntimeError(f"Could not decode/detect CSV {path}: {last}")


def new_stat(name, position):
    return {"column_name": name, "position": position, "types": Counter(), "observed": 0,
            "missing": 0, "min": None, "max": None, "categories": Counter(), "max_length": 0}


def update_stat(stat, raw, missing_tokens, profile):
    value = "" if raw is None else str(raw).strip()
    if value in missing_tokens:
        stat["missing"] += 1
        return
    stat["observed"] += 1
    kind = infer_scalar(value)
    stat["types"][kind] += 1
    stat["max_length"] = max(stat["max_length"], len(value))
    if not profile:
        return
    if kind in {"integer", "float"}:
        x = float(value)
        stat["min"] = x if stat["min"] is None else min(stat["min"], x)
        stat["max"] = x if stat["max"] is None else max(stat["max"], x)
    elif kind in {"boolean", "date", "datetime"} or len(value) <= 80:
        stat["categories"][value] += 1


def profile_csv(path, cfg):
    f, reader, encoding = open_csv(path, cfg)
    try:
        header = next(reader, None)
        if not header:
            return [], 0, encoding, "empty file"
        header = [h.strip() or f"unnamed_column_{i+1}" for i, h in enumerate(header)]
        stats = [new_stat(h, i + 1) for i, h in enumerate(header)]
        rows = 0
        for row in reader:
            rows += 1
            for i, stat in enumerate(stats):
                update_stat(stat, row[i] if i < len(row) else "", set(cfg["missing_tokens"]),
                            cfg["privacy_mode"] == "aggregate_profile")
            if rows >= int(cfg.get("sample_rows_per_file", 5000)):
                break
        return stats, rows, encoding, None
    finally:
        f.close()


def profile_xlsx(path, cfg):
    try:
        from openpyxl import load_workbook
    except ImportError:
        return [], [{"warning": "openpyxl unavailable; xlsx skipped", "file": path.name}]
    wb = load_workbook(path, read_only=True, data_only=True)
    results = []
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            continue
        stats = [new_stat(str(h).strip() if h is not None and str(h).strip() else f"unnamed_column_{i+1}", i+1)
                 for i, h in enumerate(header)]
        n = 0
        for row in rows:
            n += 1
            for i, stat in enumerate(stats):
                update_stat(stat, row[i] if i < len(row) else "", set(cfg["missing_tokens"]),
                            cfg["privacy_mode"] == "aggregate_profile")
            if n >= int(cfg.get("sample_rows_per_file", 5000)):
                break
        results.append((ws.title, stats, n))
    wb.close()
    return results, []


def merge_stats(target, incoming):
    by_pos = {s["position"]: s for s in target}
    for src in incoming:
        if src["position"] not in by_pos:
            target.append(src)
            continue
        dst = by_pos[src["position"]]
        dst["types"].update(src["types"])
        dst["observed"] += src["observed"]
        dst["missing"] += src["missing"]
        dst["max_length"] = max(dst["max_length"], src["max_length"])
        dst["categories"].update(src["categories"])
        if src["min"] is not None:
            dst["min"] = src["min"] if dst["min"] is None else min(dst["min"], src["min"])
            dst["max"] = src["max"] if dst["max"] is None else max(dst["max"], src["max"])


def serialise_column(table, stat, cfg, files_sampled, sampled_rows):
    total = stat["observed"] + stat["missing"]
    profile = cfg["privacy_mode"] == "aggregate_profile"
    cats = []
    if profile:
        threshold = int(cfg.get("rare_category_min_count", 10))
        cats = [v for v, n in stat["categories"].most_common(int(cfg.get("max_categories", 20))) if n >= threshold]
    override = cfg.get("semantic_overrides", {}).get(table, {}).get(stat["column_name"], {})
    return {
        "table_name": table,
        "column_name": stat["column_name"],
        "position": stat["position"],
        "inferred_type": override.get("type", resolved_type(stat["types"])),
        "semantic_role": override.get("semantic_role", ""),
        "unit": override.get("unit", ""),
        "files_sampled": files_sampled,
        "sampled_rows": sampled_rows,
        "missing_fraction": round(stat["missing"] / total, 4) if profile and total else None,
        "minimum": stat["min"] if profile else None,
        "maximum": stat["max"] if profile else None,
        "allowed_values": cats if profile else [],
        "max_observed_length": stat["max_length"] if profile else None,
        "notes": override.get("notes", "")
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()
    cfg = load_config(args.config)
    root = Path(cfg["input_root"]).resolve()
    out = Path(args.output).resolve()
    if not root.is_dir():
        raise SystemExit(f"Input root does not exist: {root}")
    out.mkdir(parents=True, exist_ok=True)
    grouped = {}
    warnings = []
    includes = cfg.get("include_globs", ["**/*.csv"])
    excludes = cfg.get("exclude_globs", [])
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(root)
        if not matches_any(rel, includes) or matches_any(rel, excludes):
            continue
        table = logical_table(rel, cfg)
        grouped.setdefault(table, []).append(path)
    catalog = []
    table_summary = []
    for table, paths in sorted(grouped.items()):
        merged, sampled_rows, sampled_files = [], 0, 0
        for path in paths[:int(cfg.get("max_files_per_table", 20))]:
            try:
                if path.suffix.lower() == ".csv":
                    stats, n, _, warning = profile_csv(path, cfg)
                    if warning:
                        warnings.append({"table": table, "warning": warning})
                    merge_stats(merged, stats); sampled_rows += n; sampled_files += 1
                elif path.suffix.lower() == ".xlsx":
                    sheets, ws = profile_xlsx(path, cfg); warnings.extend(ws)
                    for sheet, stats, n in sheets:
                        sheet_table = f"{table}::{sheet}"
                        for stat in stats:
                            catalog.append(serialise_column(sheet_table, stat, cfg, 1, n))
                        table_summary.append({"table_name": sheet_table, "files_found": 1, "files_sampled": 1,
                                              "sampled_rows": n, "column_count": len(stats)})
            except Exception as e:
                warnings.append({"table": table, "warning": f"{type(e).__name__}: {e}"})
        if merged:
            for stat in sorted(merged, key=lambda x: x["position"]):
                catalog.append(serialise_column(table, stat, cfg, sampled_files, sampled_rows))
            table_summary.append({"table_name": table, "files_found": len(paths), "files_sampled": sampled_files,
                                  "sampled_rows": sampled_rows, "column_count": len(merged)})
    json_path = out / "schema_catalog.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({"schema_version": 1, "privacy_mode": cfg["privacy_mode"], "columns": catalog,
                   "tables": table_summary}, f, indent=2, ensure_ascii=False)
    fields = ["table_name", "column_name", "position", "inferred_type", "semantic_role", "unit",
              "files_sampled", "sampled_rows", "missing_fraction", "minimum", "maximum", "allowed_values",
              "max_observed_length", "notes"]
    with open(out / "schema_catalog.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields); w.writeheader()
        for row in catalog:
            r = dict(row); r["allowed_values"] = json.dumps(r["allowed_values"], ensure_ascii=False); w.writerow(r)
    digest = hashlib.sha256(json_path.read_bytes()).hexdigest()
    manifest = {"extractor_version": VERSION, "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
                "privacy_mode": cfg["privacy_mode"], "table_count": len(table_summary),
                "column_count": len(catalog), "warnings": warnings, "schema_json_sha256": digest,
                "note": "No participant-level values are included in schema_only mode."}
    with open(out / "run_manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print(f"Exported {len(catalog)} columns from {len(table_summary)} tables to {out}")
    if warnings:
        print(f"Completed with {len(warnings)} warning(s); inspect run_manifest.json", file=sys.stderr)


if __name__ == "__main__":
    main()
