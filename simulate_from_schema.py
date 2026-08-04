#!/usr/bin/env python3
"""Generate non-identifying CSV data from a schema export, offline."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import math
import random
import re
from pathlib import Path


def safe_name(name):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._") or "table"


def clipped_gauss(rng, low, high):
    if low is None or high is None or low == high:
        low, high = (0.0, 100.0) if low is None or high is None else (low, high)
    mean, sd = (low + high) / 2, max((high - low) / 6, 1e-9)
    return min(high, max(low, rng.gauss(mean, sd)))


def fake_value(col, i, rng):
    name = col["column_name"].lower()
    role = (col.get("semantic_role") or "").lower()
    typ = col.get("inferred_type", "string")
    missing = col.get("missing_fraction")
    if missing is not None and rng.random() < float(missing):
        return ""
    if role in {"participant_id", "subject_id", "family_id"} or name in {"id", "participant_id", "subject_id"}:
        return f"SYN{i+1:06d}"
    allowed = col.get("allowed_values") or []
    if allowed:
        return rng.choice(allowed)
    low, high = col.get("minimum"), col.get("maximum")
    if typ == "boolean":
        return rng.choice(["0", "1"])
    if typ == "integer":
        return str(int(round(clipped_gauss(rng, low, high))))
    if typ == "float":
        return f"{clipped_gauss(rng, low, high):.4f}"
    if typ in {"date", "datetime"}:
        start = dt.date(2005, 1, 1)
        d = start + dt.timedelta(days=rng.randrange(0, 365 * 20))
        return d.isoformat() if typ == "date" else d.isoformat() + "T12:00:00"
    if any(k in name for k in ("sex", "gender")):
        return rng.choice(["Female", "Male"])
    if "code" in name:
        return f"SYN_CODE_{rng.randrange(1, 8):02d}"
    return f"synthetic_{safe_name(col['column_name']).lower()}_{rng.randrange(1, 20):02d}"


def write_schema_tables(schema, out, n, rng):
    by_table = {}
    for col in schema["columns"]:
        by_table.setdefault(col["table_name"], []).append(col)
    for table, cols in by_table.items():
        cols.sort(key=lambda c: int(c["position"]))
        path = out / f"{safe_name(table)}.csv"
        with open(path, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f); w.writerow([c["column_name"] for c in cols])
            for i in range(n):
                w.writerow([fake_value(c, i, rng) for c in cols])


def write_survival(out, n, rng):
    fields = ["participant_id", "family_id", "baseline_date", "age", "sex", "bmi", "smoker",
              "systolic_bp", "resting_hr", "mean_daily_mvpa_min", "hr_pa_response",
              "follow_up_days", "event", "event_date"]
    with open(out / "analysis_survival.csv", "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields); w.writeheader()
        for i in range(n):
            age = min(90, max(40, rng.gauss(62, 10)))
            sex = rng.choice(["Female", "Male"])
            bmi = min(45, max(17, rng.gauss(26.5, 4.5)))
            smoker = int(rng.random() < 0.15)
            sbp = min(210, max(85, rng.gauss(122 + 0.45 * (age - 50), 14)))
            rhr = min(110, max(42, rng.gauss(67 + 0.12 * (age - 60), 9)))
            mvpa = max(0, rng.gammavariate(2.2, 12))
            response = rng.gauss(0, 1)
            lp = (0.045 * (age - 60) + 0.025 * (bmi - 25) + 0.55 * smoker +
                  0.012 * (sbp - 120) + 0.018 * (rhr - 65) - 0.012 * mvpa + 0.25 * response)
            annual_rate = 0.025 * math.exp(lp)
            event_days = int(-math.log(max(rng.random(), 1e-12)) / annual_rate * 365.25)
            censor_days = rng.randint(365, int(12 * 365.25))
            follow = min(event_days, censor_days)
            event = int(event_days <= censor_days)
            baseline = dt.date(2010, 1, 1) + dt.timedelta(days=rng.randrange(0, 365 * 8))
            event_date = (baseline + dt.timedelta(days=follow)).isoformat() if event else ""
            w.writerow({"participant_id": f"SYN{i+1:06d}", "family_id": f"FAM{i//2+1:06d}",
                        "baseline_date": baseline.isoformat(), "age": f"{age:.1f}", "sex": sex,
                        "bmi": f"{bmi:.1f}", "smoker": smoker, "systolic_bp": f"{sbp:.1f}",
                        "resting_hr": f"{rhr:.1f}", "mean_daily_mvpa_min": f"{mvpa:.1f}",
                        "hr_pa_response": f"{response:.3f}", "follow_up_days": follow,
                        "event": event, "event_date": event_date})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--schema", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--rows", type=int, default=1000)
    ap.add_argument("--seed", type=int, default=20260804)
    args = ap.parse_args()
    if args.rows < 1:
        raise SystemExit("--rows must be positive")
    with open(args.schema, encoding="utf-8") as f:
        schema = json.load(f)
    out = Path(args.output); out.mkdir(parents=True, exist_ok=True)
    rng = random.Random(args.seed)
    write_schema_tables(schema, out, args.rows, rng)
    write_survival(out, args.rows, rng)
    with open(out / "synthetic_manifest.json", "w", encoding="utf-8") as f:
        json.dump({"synthetic": True, "rows_per_table": args.rows, "seed": args.seed,
                   "warning": "For software testing only; not representative of TwinsUK."}, f, indent=2)
    print(f"Synthetic data written to {out.resolve()}")


if __name__ == "__main__":
    main()
