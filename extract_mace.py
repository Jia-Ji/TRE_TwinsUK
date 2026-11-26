"""Run inside the TRE with Python, pandas >= 2.0 and openpyxl.

Edit the paths for the TRE release. Source cutoffs are inclusive. Dates are
analysed at calendar-day resolution. The phenotype detects first recorded
events, which can reflect prior disease, particularly in secondary and
outpatient diagnoses.
"""

import logging
import re
from pathlib import Path

import pandas as pd


log = logging.getLogger(__name__)

# Linux mount inside the TRE. Absolute paths also work from actiheart/scripts.
PROJECT = Path("/project/")
ACTIHEART_RELEASE = PROJECT / "Internal_Data/I1334_29.08.2025"
ACTIHEART = ACTIHEART_RELEASE / "Actiheart_Study_GPR_Processed_Files"
ACTIHEART_SUMMARY = ACTIHEART_RELEASE / "DS00016_Actiheart/Actiheart_Data.xlsx"
APC = PROJECT / "Full_Data/Admitted_Patient_Care_Data_(HES_APC).csv"
OP = PROJECT / "Full_Data/Outpatient_Care_Data_(HES_OP).csv"
MORTALITY = PROJECT / "Full_Data/Mortality_Data.csv"
CODE_LIST = PROJECT / "Full_data/CodesFollowUp_030522.csv"
OUTPUT = PROJECT / "DerivedData"
LINKED_ROSTER = None  # Path to CSV containing only eligible linked participants.
ROSTER_ID = "STUDY_ID"  # Same ID system as Study No / Person Id in this release.
# Provisional analysis assumption. A verified roster takes precedence if supplied.
ASSUME_ALL_ACTIHEART_LINKED = True
# APC and OP cutoffs are from the Linkage Data User Guide v1.0.
# The mortality cutoff is the latest date of death in the mortality spreadsheet.
SOURCE_END = {
    "APC": pd.Timestamp("2021-03-31"),
    "OP": pd.Timestamp("2021-03-31"),
    "Mortality": pd.Timestamp("2023-06-10"),
}
TIME_COLUMN = "Real_Time" 
EPOCH_COLUMN = "timepoint"

MACE = ("mi", "hf", "va", "implant")
MACE_STROKE = (*MACE, "stroke")
CATEGORY_LABELS = {"mi": "MI", "hf": "HF", "va": "VA", "implant": "implant", "af": "AF", "stroke": "stroke"}
COMPOSITES = {
    "MACE": MACE,
    "MACE_stroke": MACE_STROKE,
    "MACE_no_op": MACE,
    "MACE_stroke_no_op": MACE_STROKE,
}
UNDERLYING = {"S_UNDERLYING_COD_ICD10": "ICD10", "S_UNDERLYING_COD_ICD9": "ICD9"}
EVENT_COLUMNS = ["STUDY_ID", "event_date", "category", "code", "code_type", "source", "source_column"]


def ids(values):
    # Raw filename digits, Study No / Person Id and STUDY_ID are equivalent.
    return values.str.strip().str.replace(r"^0+(?=\d+$)", "", regex=True)


def compare_actiheart_ids(filename_ids):
    log.info("Comparing raw filename IDs with %s", ACTIHEART_SUMMARY)
    summary = pd.read_excel(ACTIHEART_SUMMARY, dtype="string", usecols=["id"])
    summary_ids = ids(summary["id"]).replace("", pd.NA).dropna()
    summary_set = set(summary_ids)
    filename_set = set(ids(pd.Series(filename_ids, dtype="string")))
    log.info("Actiheart ID check: %d summary rows with IDs (%d distinct); "
             "%d raw filenames (%d distinct IDs)",
             len(summary_ids), len(summary_set), len(filename_ids), len(filename_set))
    for label, difference in ((f"{ACTIHEART_SUMMARY.name} only", summary_set - filename_set),
                              ("Raw filenames only", filename_set - summary_set)):
        if difference:
            log.warning("%s (%d IDs): %s", label, len(difference), ", ".join(sorted(difference)))
    if summary_set == filename_set:
        log.info("Summary and raw filename ID sets match")


def dates(values):
    # UK day/month/year text and ISO timestamps (including Excel datetime cells).
    text = values.astype("string").str.strip()
    iso = text.str.match(r"\d{4}-\d{2}-\d{2}", na=False)
    parsed = pd.to_datetime(text.where(~iso), dayfirst=True, format="mixed", errors="coerce")
    return parsed.fillna(pd.to_datetime(text.where(iso), format="ISO8601", errors="coerce"))


def actiheart_baselines():
    log.info("Looking for Actiheart recordings in %s", ACTIHEART)
    paths = [path for path in sorted(ACTIHEART.iterdir())
             if re.search(r"_GPRProcessedFL(?:Data|Dtata)$", path.stem, re.I)]
    log.info("Found %d Actiheart recordings", len(paths))
    filename_ids = [path.stem.split("_", 1)[0] for path in paths]
    for path, participant_id in zip(paths, filename_ids):
        if not re.fullmatch(r"[0-9]+", participant_id):
            raise ValueError(f"Expected a numeric Actiheart ID before the first underscore in {path.name}")
    compare_actiheart_ids(filename_ids)
    recordings = []
    for number, (path, participant_id) in enumerate(zip(paths, filename_ids), start=1):
        readers = {".csv": pd.read_csv, ".xlsx": pd.read_excel}
        if path.suffix.lower() not in readers:
            raise ValueError(f"Check the recording file extension for {path.name}")
        raw = readers[path.suffix.lower()](path, usecols=[TIME_COLUMN, EPOCH_COLUMN])
        # Count distinct timestamped epochs from timepoint zero onward. Negative
        # timepoints in the example are before recording start. No wear threshold
        # was supplied, so any recording with a usable timestamp is eligible.
        epochs = dates(raw[TIME_COLUMN][pd.to_numeric(raw[EPOCH_COLUMN], errors="coerce").ge(0)]).dropna()
        recordings.append((participant_id, epochs.min(), epochs.nunique()))
        if number % 50 == 0 or number == len(paths):
            log.info("Actiheart: processed %d/%d recordings", number, len(paths))
    if not recordings:
        raise ValueError(f"No processed Actiheart recordings found in {ACTIHEART}")
    cohort = pd.DataFrame(recordings, columns=["STUDY_ID", "baseline_date", "usable_epochs"])
    cohort["STUDY_ID"] = ids(cohort["STUDY_ID"].astype("string"))
    # Keep the recording with most usable epochs. Break ties by earliest start.
    cohort = cohort.sort_values(["usable_epochs", "baseline_date"], ascending=[False, True])
    cohort = cohort.drop_duplicates("STUDY_ID").set_index("STUDY_ID")
    cohort["baseline_date"] = cohort["baseline_date"].dt.normalize()
    cohort["linked"] = pd.Series(pd.NA, index=cohort.index, dtype="boolean")
    cohort["linkage_assumed"] = LINKED_ROSTER is None and ASSUME_ALL_ACTIHEART_LINKED
    if LINKED_ROSTER is not None:
        roster = pd.read_csv(LINKED_ROSTER, dtype="string", usecols=[ROSTER_ID])
        cohort["linked"] = cohort.index.isin(ids(roster[ROSTER_ID]).dropna())
    elif ASSUME_ALL_ACTIHEART_LINKED:
        cohort["linked"] = True
    return cohort.sort_index()


def code_lookup():
    codes = pd.read_csv(CODE_LIST, dtype="string", usecols=["code", "cat", "type"]).dropna()
    codes["code"] = normalise_codes(codes["code"])
    codes["cat"] = codes["cat"].str.strip().str.lower()
    codes["type"] = codes["type"].str.strip().str.upper()
    codes = codes[codes["cat"].isin((*CATEGORY_LABELS, "cv"))
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
    log.info("%s: opening %s", source, path)
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
    rows_read = cohort_rows = 0
    needed = {"STUDY_ID", *fields, *date_fields}
    for data in pd.read_csv(path, usecols=lambda column: column in needed, dtype="string", chunksize=50000):
        rows_read += len(data)
        data["STUDY_ID"] = ids(data["STUDY_ID"])
        data = data[data["STUDY_ID"].isin(cohort.index)]
        cohort_rows += len(data)
        log.info("%s: read %s rows (%s from the Actiheart cohort)",
                 source, f"{rows_read:,}", f"{cohort_rows:,}")
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
            for category in CATEGORY_LABELS:
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
    log.info("%s: finished matching %s cohort rows; %s code-position matches lacked a usable date",
             source, f"{cohort_rows:,}", f"{missing_dates:,}")
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
    # Combined outcomes require observation from all three contributing sources.
    mace_end = min(SOURCE_END.values())
    mortality_end = SOURCE_END["Mortality"]
    linked = out["linked"].fillna(False)
    death_date = out["death_date"]
    eligible = (linked & ~out["death_date_missing"]
                & (death_date.isna() | death_date.gt(baseline)))
    mace_observable = eligible & baseline.lt(mace_end)
    mortality_observable = eligible & baseline.lt(mortality_end)
    mace_censor = death_date.where(death_date.le(mace_end), mace_end)
    mortality_censor = death_date.where(death_date.le(mortality_end), mortality_end)

    timed = events.join(out[["baseline_date", "death_date", "linked"]], on="STUDY_ID")
    timed = timed[timed["linked"].fillna(True) & timed["baseline_date"].notna()]
    timed = timed[timed["event_date"].le(timed["source"].map(SOURCE_END))]
    timed = timed[timed["death_date"].isna() | timed["event_date"].le(timed["death_date"])]
    history = timed[timed["event_date"].le(timed["baseline_date"])]
    incident = timed[timed["event_date"].gt(timed["baseline_date"]) & timed["event_date"].le(mace_end)]
    incident = incident.sort_values(["event_date", "source", "source_column", "code"])

    endpoints = {label: (cat,) for cat, label in CATEGORY_LABELS.items()}
    endpoints.update(COMPOSITES)
    for name, categories in endpoints.items():
        incident_match = incident["category"].isin(categories)
        history_match = history["category"].isin(categories)
        if name.endswith("_no_op"):
            # Apply the code-system filter before choosing a first event or flagging history.
            incident_match &= incident["code_type"].ne("OPCS4")
            history_match &= history["code_type"].ne("OPCS4")
        first = incident[incident_match].drop_duplicates("STUDY_ID").set_index("STUDY_ID")
        out[f"{name}_date"] = first["event_date"]
        out[f"{name}_code"] = first["code"]
        out[f"{name}_source"] = first["source"]
        if name in COMPOSITES:
            prior = history.loc[history_match, "STUDY_ID"]
            out[f"{name}_history"] = pd.Series(out.index.isin(prior), index=out.index, dtype="boolean")
            out.loc[~linked | baseline.isna(), f"{name}_history"] = pd.NA
            out[f"{name}_event"] = out[f"{name}_date"].notna().astype("Int64").where(mace_observable)
            followup_days = (out[f"{name}_date"].fillna(mace_censor) - baseline).dt.days
            out[f"{name}_followup_days"] = followup_days.astype("Int64").where(mace_observable)

    # Separate mortality endpoints. Only underlying CV codes count toward CVM.
    # A known non-CV death ends CVM follow-up. Unknown causes remain unknown.
    post_baseline_death = death_date.gt(baseline) & death_date.le(mortality_end) & out["linked"].fillna(True)
    cv_death = post_baseline_death & out["FU_CVM_code"].notna()
    known_cause = out.pop("cause_known").astype("boolean").fillna(False)
    for name, event, known in (("FU_ACM", post_baseline_death, mortality_observable),
                               ("FU_CVM", cv_death, mortality_observable & (~post_baseline_death | known_cause))):
        out[name] = event.astype("Int64").where(known)
        out[f"{name}_date"] = death_date.where(event)
        out[f"{name}_followup_days"] = (mortality_censor - baseline).dt.days.astype("Int64").where(known)
    out["FU_CVM_code"] = out["FU_CVM_code"].where(cv_death)
    out["CVM_non_CV_death"] = (post_baseline_death & ~cv_death & known_cause).where(mortality_observable)
    out["CVM_unknown_cause_death"] = (post_baseline_death & ~known_cause).where(mortality_observable)

    for broader, narrower in (("MACE_stroke", "MACE"), ("MACE_stroke_no_op", "MACE_no_op"),
                              ("MACE", "MACE_no_op"), ("MACE_stroke", "MACE_stroke_no_op")):
        has_event = out[f"{narrower}_date"].notna()
        assert out.loc[has_event, f"{broader}_date"].le(out.loc[has_event, f"{narrower}_date"]).all()
    assert out.index.is_unique
    assert out.filter(like="_followup_days").ge(0).fillna(True).all().all()
    return out


def main():
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    log.info("Starting TwinsUK outcome extraction")
    if LINKED_ROSTER is None and ASSUME_ALL_ACTIHEART_LINKED:
        log.warning("Assuming all raw Actiheart participants are eligible for HES and mortality linkage.")
    cohort = actiheart_baselines()
    linked_count = int(cohort["linked"].sum()) if cohort["linked"].notna().any() else "unknown"
    log.info("Actiheart cohort: %d participants, %s eligible for linkage, %d missing baseline",
             len(cohort), linked_count, cohort["baseline_date"].isna().sum())
    log.info("Reading outcome codes from %s", CODE_LIST)
    lookup = code_lookup()
    event_parts, deaths = [], []
    for path, source in ((APC, "APC"), (OP, "OP"), (MORTALITY, "Mortality")):
        tables, death_records = extract(path, source, cohort, lookup)
        event_parts.extend(tables)
        deaths.extend(death_records)
    events = (pd.concat(event_parts, ignore_index=True) if event_parts else
              pd.DataFrame(columns=EVENT_COLUMNS).astype({"event_date": "datetime64[ns]"}))
    events = events.drop_duplicates().sort_values(["STUDY_ID", "event_date", "category"])
    # Keep matches outside coverage in the audit, with an explicit flag.
    events["within_source_coverage"] = events["event_date"].le(events["source"].map(SOURCE_END))
    log.info("Summarising %d distinct matches for %d Actiheart participants", len(events), len(cohort))
    outcomes = summarise(cohort, events, deaths)

    log.info("Writing mace_events.csv and mace_outcomes.csv to %s", OUTPUT)
    OUTPUT.mkdir(parents=True, exist_ok=True)
    events.to_csv(OUTPUT / "mace_events.csv", index=False)
    outcomes.to_csv(OUTPUT / "mace_outcomes.csv")

    log.info("Participants with a mortality record and no usable death date: %d",
             outcomes["death_date_missing"].sum())
    counts = events.groupby(["source", "category"]).agg(
        matches=("code", "size"), participants=("STUDY_ID", "nunique"),
    )
    log.info("Audit matches by source and category:\n%s", counts)
    for name in COMPOSITES:
        log.info("%s: %d with confirmed history, %d with a post-baseline date",
                 name, outcomes[f"{name}_history"].sum(), outcomes[f"{name}_date"].count())
    event_columns = {name: f"{name}_event" for name in COMPOSITES}
    event_columns.update({"all-cause mortality": "FU_ACM", "cardiovascular mortality": "FU_CVM"})
    for label, column in event_columns.items():
        log.info("%s: %d events among %d participants with final follow-up",
                 label, outcomes[column].sum(), outcomes[column].count())
    if LINKED_ROSTER is None and not ASSUME_ALL_ACTIHEART_LINKED:
        log.warning("Provisional output: set the verified linkage roster for final survival indicators.")
    log.info("Extraction complete: wrote %d audit rows and %d participant rows", len(events), len(outcomes))


if __name__ == "__main__":
    main()
