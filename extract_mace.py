"""Run inside the TRE with Python and pandas >= 2.0 (openpyxl for XLSX recordings).

Edit the paths and set the common coverage end only after verifying coverage
for APC, OP and mortality. Dates without times are analysed at calendar-day
resolution. The phenotype detects first recorded events, which can reflect
prior disease, particularly in secondary and outpatient diagnoses.
"""

import re
from pathlib import Path

import pandas as pd


PROJECT = Path(r"\\192.168.0.10\storage\project")
ACTIHEART = PROJECT / "Internal_Data/I1334_29.08.2025/Actiheart_Study_GPR_Processed_Files"
APC = PROJECT / "Full_Data/Admitted_Patient_Care_Data_(HES_APC).csv"
OP = PROJECT / "Full_Data/Outpatient_Care_Data_(HES_OP).csv"
MORTALITY = PROJECT / "Full_Data/Mortality_Data.csv"
CODE_LIST = PROJECT / "CodesFollowUp_030522.csv"
OUTPUT = PROJECT / "DerivedData"
LINKED_ROSTER = None  # Path to CSV containing only eligible linked participants.
ROSTER_ID = "STUDY_ID"  # Same ID system as Study No / Person Id in this release.
COVERAGE_END = None  # Verified common inclusive end date, as "YYYY-MM-DD".
TIME_COLUMN = "Real_Time"  # Header shown in the raw recording screenshot.
EPOCH_COLUMN = "timepoint"

MACE = ("mi", "hf", "va", "implant", "af")
MACE_STROKE = (*MACE, "stroke")
CATEGORY_LABELS = dict(zip(MACE_STROKE, ("MI", "HF", "VA", "implant", "AF", "stroke")))
UNDERLYING = {"S_UNDERLYING_COD_ICD10": "ICD10", "S_UNDERLYING_COD_ICD9": "ICD9"}
EVENT_COLUMNS = ["STUDY_ID", "event_date", "category", "code", "code_type", "source", "source_column"]


def ids(values):
    # Raw filename digits, Study No / Person Id and STUDY_ID are equivalent.
    return values.str.strip().str.replace(r"^0+(?=\d+$)", "", regex=True)


def dates(values):
    # UK day/month/year text and ISO timestamps (including Excel datetime cells).
    text = values.astype("string").str.strip()
    iso = text.str.match(r"\d{4}-\d{2}-\d{2}", na=False)
    parsed = pd.to_datetime(text.where(~iso), dayfirst=True, format="mixed", errors="coerce")
    return parsed.fillna(pd.to_datetime(text.where(iso), format="ISO8601", errors="coerce"))


def actiheart_baselines():
    recordings = []
    for path in sorted(ACTIHEART.iterdir()):
        if not re.search(r"_GPRProcessedFL(?:Data|Dtata)$", path.stem, re.I):
            continue
        match = re.search(r"#(\d+)[^#]*#", path.stem)
        if not match:
            raise ValueError(f"Cannot find the Actiheart ID in {path.name}")
        readers = {".csv": pd.read_csv, ".xlsx": pd.read_excel}
        if path.suffix.lower() not in readers:
            raise ValueError(f"Check the recording file extension for {path.name}")
        raw = readers[path.suffix.lower()](path, usecols=[TIME_COLUMN, EPOCH_COLUMN])
        # Count distinct timestamped epochs from timepoint zero onward. Negative
        # timepoints in the example are before recording start. No wear threshold
        # was supplied, so any recording with a usable timestamp is eligible.
        epochs = dates(raw[TIME_COLUMN][pd.to_numeric(raw[EPOCH_COLUMN], errors="coerce").ge(0)]).dropna()
        recordings.append((match[1], epochs.min(), epochs.nunique()))
    if not recordings:
        raise ValueError(f"No processed Actiheart recordings found in {ACTIHEART}")
    cohort = pd.DataFrame(recordings, columns=["STUDY_ID", "baseline_date", "usable_epochs"])
    cohort["STUDY_ID"] = ids(cohort["STUDY_ID"].astype("string"))
    # Keep the recording with most usable epochs. Break ties by earliest start.
    cohort = cohort.sort_values(["usable_epochs", "baseline_date"], ascending=[False, True])
    cohort = cohort.drop_duplicates("STUDY_ID").set_index("STUDY_ID")
    cohort["baseline_date"] = cohort["baseline_date"].dt.normalize()
    cohort["linked"] = pd.Series(pd.NA, index=cohort.index, dtype="boolean")
    if LINKED_ROSTER is not None:
        roster = pd.read_csv(LINKED_ROSTER, dtype="string", usecols=[ROSTER_ID])
        cohort["linked"] = cohort.index.isin(ids(roster[ROSTER_ID]).dropna())
    return cohort.sort_index()


def code_lookup():
    codes = pd.read_csv(CODE_LIST, dtype="string", usecols=["code", "cat", "type"]).dropna()
    codes["code"] = normalise_codes(codes["code"])
    codes["cat"] = codes["cat"].str.strip().str.lower()
    codes["type"] = codes["type"].str.strip().str.upper()
    codes = codes[codes["cat"].isin((*MACE_STROKE, "cv"))
                  & codes["type"].isin(["ICD9", "ICD10", "OPCS4"])].drop_duplicates()
    lookup = {kind: {} for kind in ("ICD9", "ICD10", "OPCS4")}
    for (kind, category), rows in codes.groupby(["type", "cat"]):
        lookup[kind][category] = set(rows["code"])
    return lookup


def normalise_codes(values):
    return values.str.strip().str.upper().str.replace(".", "", regex=False)


def matches(values, kind, categories):
    codes = normalise_codes(values)
    if kind == "ICD9":
        return codes, codes.isin(categories)
    codes = codes.str[:4]
    parents = {code for code in categories if len(code) == 3}
    return codes, codes.isin(categories) | codes.str[:3].isin(parents)


def extract(path, source, cohort, lookup):
    headers = pd.read_csv(path, nrows=0).columns.tolist()
    if source == "Mortality":
        fields = {column: UNDERLYING[column] for column in UNDERLYING if column in headers}
        for prefix, kind in (("S_COD_CODE", "ICD10"), ("ICD9_ORIG_MENTION", "ICD9")):
            fields.update({column: kind for column in headers if re.fullmatch(prefix + r"_\d+", column)})
        date_column = "REG_DATE_OF_DEATH"
    else:
        fields = {column: "ICD10" for column in headers if re.fullmatch(r"DIAG_\d+", column)}
        fields.update({column: "OPCS4" for column in headers if re.fullmatch(r"OPERTN_\d+", column)})
        date_column = "ADMIDATE" if source == "APC" else "APPTDATE"
    if not fields or "STUDY_ID" not in headers:
        raise ValueError(f"Check the {source} ID and code headers in {path.name}")
    if source == "Mortality" and not (fields.keys() & UNDERLYING.keys()):
        raise ValueError("Check the underlying cause-of-death headers for cardiovascular mortality")
    date_fields = {date_column}
    if source == "APC":
        date_fields.update(column for column in headers
                           if column == "EPISTART" or re.fullmatch(r"OPDATE_\d+", column))
    if not (date_fields & set(headers)):
        raise ValueError(f"Check the event date headers in {path.name}")

    events, deaths = [], []
    missing_dates = 0
    needed = {"STUDY_ID", *fields, *date_fields}
    for data in pd.read_csv(path, usecols=lambda column: column in needed, dtype="string", chunksize=50000):
        data["STUDY_ID"] = ids(data["STUDY_ID"])
        data = data[data["STUDY_ID"].isin(cohort.index)]
        if data.empty:
            continue
        event_dates = {column: dates(data[column]).dt.normalize() for column in date_fields if column in data}
        default_date = event_dates.get(date_column, pd.Series(pd.NaT, index=data.index, dtype="datetime64[ns]"))
        if "EPISTART" in event_dates:
            default_date = default_date.fillna(event_dates["EPISTART"])
        if source == "Mortality":
            death = data[["STUDY_ID"]].copy()
            death["death_date"] = default_date
            death["cause_known"] = False
            death["FU_CVM_code"] = pd.Series(pd.NA, index=data.index, dtype="string")
            for column, kind in UNDERLYING.items():
                if column in data:
                    code, cv = matches(data[column], kind, lookup[kind].get("cv", set()))
                    valid_pattern = r"[A-Z]\d{2}[A-Z0-9]?" if kind == "ICD10" else r"(?:\d{3,5}|[EV]\d{2,4})"
                    death["cause_known"] |= code.str.fullmatch(valid_pattern, na=False)
                    death["FU_CVM_code"] = death["FU_CVM_code"].fillna(code.where(cv))
            deaths.append(death)

        for column, kind in fields.items():
            date = default_date
            if source == "APC" and kind == "OPCS4":
                operation_date = event_dates.get(column.replace("OPERTN_", "OPDATE_", 1))
                if operation_date is not None:
                    date = operation_date.fillna(default_date)
            for category in MACE_STROKE:
                listed = lookup[kind].get(category)
                if not listed:
                    continue
                code, hit = matches(data[column], kind, listed)
                missing_dates += int((hit & date.isna()).sum())
                keep = hit & date.notna()
                if keep.any():
                    events.append(pd.DataFrame({
                        "STUDY_ID": data.loc[keep, "STUDY_ID"], "event_date": date[keep],
                        "category": category, "code": code[keep], "code_type": kind,
                        "source": source, "source_column": column,
                    }))
    print(f"{source}: {missing_dates} qualifying code-position matches missing a usable date")
    return events, deaths


def summarise(cohort, events, death_records):
    if death_records:
        all_deaths = pd.concat(death_records)
        death = all_deaths.dropna(subset=["death_date"])
        if death.groupby("STUDY_ID")["death_date"].nunique().gt(1).any():
            raise ValueError("Conflicting dates of death for the same STUDY_ID")
        death = death.groupby("STUDY_ID").agg(
            death_date=("death_date", "first"), cause_known=("cause_known", "max"),
            FU_CVM_code=("FU_CVM_code", "first"),
        )
        undated_deaths = set(all_deaths["STUDY_ID"]) - set(death.index)
    else:
        death = pd.DataFrame({"death_date": pd.Series(dtype="datetime64[ns]"),
                              "cause_known": pd.Series(dtype="boolean"),
                              "FU_CVM_code": pd.Series(dtype="string")})
        undated_deaths = set()
    out = cohort.join(death)
    out["death_date_missing"] = out.index.isin(undated_deaths)
    baseline = out["baseline_date"]
    end = pd.Timestamp(COVERAGE_END) if COVERAGE_END else pd.NaT
    linked = out["linked"].fillna(False)
    death_date = out["death_date"]
    observable = (linked & baseline.lt(end) & ~out["death_date_missing"]
                  & (death_date.isna() | death_date.gt(baseline)))
    censor = pd.concat([death_date, pd.Series(end, index=out.index)], axis=1).min(axis=1)

    timed = events.join(out[["baseline_date", "death_date", "linked"]], on="STUDY_ID")
    timed = timed[timed["linked"].fillna(True) & timed["baseline_date"].notna()]
    timed = timed[timed["death_date"].isna() | timed["event_date"].le(timed["death_date"])]
    history = timed[timed["event_date"].le(timed["baseline_date"])]
    incident = timed[timed["event_date"].gt(timed["baseline_date"])]
    if pd.notna(end):
        incident = incident[incident["event_date"].le(end)]
    incident = incident.sort_values(["event_date", "source", "source_column", "code"])

    endpoints = {label: (cat,) for cat, label in CATEGORY_LABELS.items()}
    endpoints.update({"MACE": MACE, "MACE_stroke": MACE_STROKE})
    for name, categories in endpoints.items():
        first = incident[incident["category"].isin(categories)].drop_duplicates("STUDY_ID").set_index("STUDY_ID")
        out[f"{name}_date"] = first["event_date"]
        out[f"{name}_code"] = first["code"]
        out[f"{name}_source"] = first["source"]
        if name in ("MACE", "MACE_stroke"):
            prior = history.loc[history["category"].isin(categories), "STUDY_ID"]
            out[f"{name}_history"] = pd.Series(out.index.isin(prior), index=out.index, dtype="boolean")
            out.loc[~linked | baseline.isna(), f"{name}_history"] = pd.NA
            out[f"{name}_event"] = out[f"{name}_date"].notna().astype("Int64").where(observable)
            out[f"{name}_followup_days"] = (out[f"{name}_date"].fillna(censor) - baseline).dt.days.astype("Int64").where(observable)

    # Separate mortality endpoints. Only underlying CV codes count toward CVM.
    # A known non-CV death ends CVM follow-up. Unknown causes remain unknown.
    post_baseline_death = death_date.gt(baseline) & out["linked"].fillna(True)
    if pd.notna(end):
        post_baseline_death &= death_date.le(end)
    cv_death = post_baseline_death & out["FU_CVM_code"].notna()
    known_cause = out.pop("cause_known").astype("boolean").fillna(False)
    for name, event, known in (("FU_ACM", post_baseline_death, observable),
                               ("FU_CVM", cv_death, observable & (~post_baseline_death | known_cause))):
        out[name] = event.astype("Int64").where(known)
        out[f"{name}_date"] = death_date.where(event)
        out[f"{name}_followup_days"] = (censor - baseline).dt.days.astype("Int64").where(known)
    out["FU_CVM_code"] = out["FU_CVM_code"].where(cv_death)
    out["CVM_non_CV_death"] = (post_baseline_death & ~cv_death & known_cause).where(observable)
    out["CVM_unknown_cause_death"] = (post_baseline_death & ~known_cause).where(observable)

    has_mace = out["MACE_date"].notna()
    assert out.loc[has_mace, "MACE_stroke_date"].le(out.loc[has_mace, "MACE_date"]).all()
    assert out.index.is_unique
    assert out.filter(like="_followup_days").ge(0).fillna(True).all().all()
    return out


def main():
    cohort = actiheart_baselines()
    lookup = code_lookup()
    event_parts, deaths = [], []
    for path, source in ((APC, "APC"), (OP, "OP"), (MORTALITY, "Mortality")):
        tables, death_records = extract(path, source, cohort, lookup)
        event_parts.extend(tables)
        deaths.extend(death_records)
    events = (pd.concat(event_parts, ignore_index=True) if event_parts else
              pd.DataFrame(columns=EVENT_COLUMNS).astype({"event_date": "datetime64[ns]"}))
    events = events.drop_duplicates().sort_values(["STUDY_ID", "event_date", "category"])
    outcomes = summarise(cohort, events, deaths)

    OUTPUT.mkdir(parents=True, exist_ok=True)
    events.to_csv(OUTPUT / "mace_events.csv", index=False)
    outcomes.to_csv(OUTPUT / "mace_outcomes.csv")

    linked_count = int(cohort["linked"].sum()) if LINKED_ROSTER is not None else "unknown"
    print(f"Actiheart participants: {len(cohort)} | linked: {linked_count} | missing baseline: {cohort['baseline_date'].isna().sum()}")
    print(f"Participants with a mortality record and no usable death date: {outcomes['death_date_missing'].sum()}")
    print(events.groupby(["source", "category"]).agg(matches=("code", "size"), participants=("STUDY_ID", "nunique")))
    for name in ("MACE", "MACE_stroke"):
        print(f"{name}: {outcomes[f'{name}_history'].sum()} with confirmed history, "
              f"{outcomes[f'{name}_date'].count()} with a post-baseline date")
    for label, column in (("MACE", "MACE_event"), ("MACE_stroke", "MACE_stroke_event"),
                          ("all-cause mortality", "FU_ACM"), ("cardiovascular mortality", "FU_CVM")):
        print(f"{label}: {outcomes[column].sum()} events among {outcomes[column].count()} with final follow-up")
    if LINKED_ROSTER is None or COVERAGE_END is None:
        print("Provisional output: set the verified linkage roster and coverage end for final survival indicators.")


if __name__ == "__main__":
    main()
