# -*- coding: utf-8 -*-
"""
rkei_parser.py

Refactored from Colab notebook into an importable module.
Exposes: process_files(file_paths: List[str]) -> bytes
"""

import os
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import List

import pandas as pd

# -------------------------
# Constants
# -------------------------
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

SPRE_CODES = {"STR", "PPL", "IIF", "CEI"}
PARTNER_CODES = {"ACD", "IND", "CUL", "COM", "PUB", "PRO", "NON"}
STAGE_CODES = {"PLN", "DEV", "EXT", "LIV", "CMP", "EVD"}

CODE_MEANINGS = {
    "STR": "Strategic research output",
    "PPL": "People / research culture",
    "IIF": "Income / funding",
    "CEI": "Civic / external impact",
    "PLN": "Planning stage",
    "DEV": "In development",
    "EXT": "Externally submitted",
    "LIV": "Live activity",
    "CMP": "Completed",
    "EVD": "Evidence of impact",
    "ACD": "Academic partner",
    "IND": "Industry partner",
    "CUL": "Cultural organisation",
    "COM": "Community group",
    "PUB": "Public sector",
    "PRO": "Professional body",
    "NON": "Non-profit",
}

# Table indexes (template-based; adjust if your template differs)
TABLE_IDX = {
    "staff": 3,
    "priorities": 5,
    "bids": 9,
    "events": 11,
    "engagement": 12,
    "impact": 13,
}

# -------------------------
# XML helpers
# -------------------------
def get_root(path):
    with zipfile.ZipFile(path, "r") as z:
        return ET.fromstring(z.read("word/document.xml"))


def text(elem):
    return " ".join([t.text for t in elem.findall(".//w:t", NS) if t.text]).strip()


def dropdowns(tc):
    vals = []
    for sdt in tc.findall(".//w:sdt", NS):
        val = text(sdt).strip()
        val = re.sub(r".*:\s*", "", val)
        if val:
            vals.append(val)
    return vals


def first_dropdown_or_text(tc):
    vals = dropdowns(tc)
    if vals:
        return vals[0]
    return text(tc)


def first(tc, allowed):
    for v in dropdowns(tc):
        if v in allowed:
            return v
    return ""

# -------------------------
# Date detection & normalization
# -------------------------
DATE_PATTERNS = [
    r"^\d{1,2}/\d{1,2}/\d{2,4}$",
    r"^\d{4}-\d{1,2}-\d{1,2}$",
    r"^[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}$",
]


def looks_like_date(s):
    if not s or not isinstance(s, str):
        return False
    s = s.strip()
    for p in DATE_PATTERNS:
        if re.match(p, s):
            return True
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%b %d, %Y", "%B %d, %Y"):
        try:
            datetime.strptime(s, fmt)
            return True
        except Exception:
            pass
    return False


def normalize_staff_vals(staff_vals):
    vals = [v.strip() if isinstance(v, str) else "" for v in staff_vals]
    review_date = ""
    date_indices = [i for i, v in enumerate(vals) if looks_like_date(v)]
    if date_indices:
        idx = date_indices[-1]
        review_date = vals[idx]
        vals.pop(idx)
    while len(vals) < 5:
        vals.append("")
    return {
        "staff_name": vals[0],
        "position": vals[1],
        "department": vals[2],
        "pathway": vals[3],
        "uoa": vals[4],
        "review_date": review_date,
    }

# -------------------------
# Parse a single document
# -------------------------
def parse_doc(path):
    root = get_root(path)
    tables = root.findall(".//w:tbl", NS)
    fname = os.path.basename(path)

    def cells(r):
        return r.findall(".//w:tc", NS)

    def rows(t):
        return t.findall(".//w:tr", NS)

    staff_row = None
    try:
        staff_tbl = tables[TABLE_IDX["staff"]]
        staff_rows = rows(staff_tbl)
        if len(staff_rows) > 1:
            staff_row = staff_rows[1]
    except Exception:
        staff_row = None

    staff_vals = []
    if staff_row is not None:
        staff_cells = cells(staff_row)
        staff_vals = [first_dropdown_or_text(c) for c in staff_cells]

    meta_staff = normalize_staff_vals(staff_vals)
    meta = {"file_name": fname, **meta_staff}

    data = []
    # Priorities
    try:
        for r in rows(tables[TABLE_IDX["priorities"]])[1:6]:
            c = cells(r)
            if len(c) < 7:
                continue
            stages = [v for v in dropdowns(c[4]) if v in STAGE_CODES]
            data.append(
                {
                    **meta,
                    "section": "Priorities",
                    "row_id": text(c[0]),
                    "entry": text(c[1]),
                    "spre_code": first(c[3], SPRE_CODES),
                    "baseline": stages[0] if len(stages) > 0 else "",
                    "target": stages[1] if len(stages) > 1 else "",
                }
            )
    except Exception:
        pass
    # Bids
    try:
        for r in rows(tables[TABLE_IDX["bids"]])[1:5]:
            c = cells(r)
            if len(c) < 5:
                continue
            data.append(
                {
                    **meta,
                    "section": "Bids",
                    "row_id": text(c[0]),
                    "entry": text(c[1]),
                    "stage": first(c[4], STAGE_CODES),
                }
            )
    except Exception:
        pass
    # Events
    try:
        for r in rows(tables[TABLE_IDX["events"]])[1:4]:
            c = cells(r)
            if len(c) < 5:
                continue
            data.append(
                {
                    **meta,
                    "section": "Events",
                    "row_id": text(c[0]),
                    "entry": text(c[1]),
                    "partner": first(c[4], PARTNER_CODES),
                }
            )
    except Exception:
        pass
    # Engagement
    try:
        for r in rows(tables[TABLE_IDX["engagement"]])[1:4]:
            c = cells(r)
            if len(c) < 3:
                continue
            data.append(
                {
                    **meta,
                    "section": "Engagement",
                    "row_id": text(c[0]),
                    "entry": text(c[1]),
                    "partner": first(c[2], PARTNER_CODES),
                }
            )
    except Exception:
        pass
    # Impact
    try:
        for r in rows(tables[TABLE_IDX["impact"]])[1:4]:
            c = cells(r)
            if len(c) < 4:
                continue
            data.append(
                {
                    **meta,
                    "section": "Impact",
                    "row_id": text(c[0]),
                    "entry": text(c[1]),
                    "stage": first(c[3], STAGE_CODES),
                }
            )
    except Exception:
        pass

    return data


# ============================================================
# PUBLIC API
# ============================================================
def process_files(file_paths: List[str]) -> bytes:
    """
    Accept a list of filesystem paths to .docx RKEI forms,
    parse them, and return an Excel workbook as bytes.

    Sheets: master_rows, file_level_extraction, pivot_pathway_uoa,
            summary_counts, counts_by_file, distinct_staff_counts,
            codes_by_uoa, codes_by_pathway.
    """
    records = []
    for fpath in file_paths:
        try:
            records += parse_doc(fpath)
        except Exception as e:
            print(f"Error parsing {os.path.basename(fpath)}: {e}")

    df = pd.DataFrame(records)

    # ---- Code lists & summary ----
    codes = []
    for _, r in df.iterrows():
        if r.get("spre_code"):
            codes.append(("SPRE", r["spre_code"]))
        if r.get("stage"):
            codes.append(("STAGE", r["stage"]))
        if r.get("partner"):
            codes.append(("PARTNER", r["partner"]))
        if r.get("baseline"):
            codes.append(("STAGE", r["baseline"]))
        if r.get("target"):
            codes.append(("STAGE", r["target"]))

    codes_df = pd.DataFrame(codes, columns=["family", "code"])
    summary = pd.DataFrame(columns=["family", "code", "count", "meaning", "percent"])
    if not codes_df.empty:
        summary = codes_df.groupby(["family", "code"]).size().reset_index(name="count")
        summary["meaning"] = summary["code"].map(CODE_MEANINGS)
        summary["percent"] = summary.groupby("family")["count"].transform(
            lambda x: (x / x.sum() * 100).round(1)
        )
        summary = summary.sort_values(["family", "count"], ascending=[True, False])

    # ---- File-level pivot Pathway x UoA ----
    forms = pd.DataFrame()
    if not df.empty:
        forms = df[["file_name", "pathway", "uoa"]].drop_duplicates()
    pivot_by_pathway_uoa = pd.DataFrame()
    if not forms.empty:
        pivot_by_pathway_uoa = (
            forms.pivot_table(
                index="pathway",
                columns="uoa",
                values="file_name",
                aggfunc="nunique",
                fill_value=0,
            )
            .sort_index()
            .sort_index(axis=1)
        )
        pivot_by_pathway_uoa["Total_by_pathway"] = pivot_by_pathway_uoa.sum(axis=1)
        pivot_by_pathway_uoa.loc["Total_by_uoa"] = pivot_by_pathway_uoa.sum(axis=0)

    # ---- Other aggregations ----
    counts_by_file = (
        df.groupby("file_name").size().reset_index(name="entries")
        if not df.empty
        else pd.DataFrame(columns=["file_name", "entries"])
    )
    distinct_staff = (
        df.groupby("staff_name").size().reset_index(name="entries")
        if not df.empty
        else pd.DataFrame(columns=["staff_name", "entries"])
    )

    # Build long-form codes tied to uoa/pathway
    long_codes = []
    for _, row in df.iterrows():
        m = {
            "file_name": row.get("file_name", ""),
            "staff_name": row.get("staff_name", ""),
            "uoa": row.get("uoa", ""),
            "pathway": row.get("pathway", ""),
        }
        if row.get("spre_code"):
            long_codes.append({**m, "family": "SPRE", "code": row["spre_code"]})
        if row.get("stage"):
            long_codes.append({**m, "family": "STAGE", "code": row["stage"]})
        if row.get("partner"):
            long_codes.append({**m, "family": "PARTNER", "code": row["partner"]})
        if row.get("baseline"):
            long_codes.append({**m, "family": "STAGE", "code": row["baseline"]})
        if row.get("target"):
            long_codes.append({**m, "family": "STAGE", "code": row["target"]})

    long_codes_df = pd.DataFrame(long_codes)
    if not long_codes_df.empty:
        codes_by_uoa = (
            long_codes_df.groupby(["uoa", "family", "code"])
            .size()
            .reset_index(name="count")
            .sort_values(["uoa", "family", "count"], ascending=[True, True, False])
        )
        codes_by_pathway = (
            long_codes_df.groupby(["pathway", "family", "code"])
            .size()
            .reset_index(name="count")
            .sort_values(
                ["pathway", "family", "count"], ascending=[True, True, False]
            )
        )
    else:
        codes_by_uoa = pd.DataFrame(columns=["uoa", "family", "code", "count"])
        codes_by_pathway = pd.DataFrame(
            columns=["pathway", "family", "code", "count"]
        )

    # ---- Diagnostics ----
    diagnostics = (
        forms.copy()
        if not forms.empty
        else pd.DataFrame(columns=["file_name", "pathway", "uoa"])
    )
    if not diagnostics.empty:
        diagnostics["pathway_missing"] = (
            diagnostics["pathway"].fillna("").apply(lambda x: x == "")
        )
        diagnostics["uoa_missing"] = (
            diagnostics["uoa"].fillna("").apply(lambda x: x == "")
        )

    # ---- Write Excel to in-memory buffer ----
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="summary_counts", index=False)
        counts_by_file.to_excel(writer, sheet_name="counts_by_file", index=False)
        distinct_staff.to_excel(
            writer, sheet_name="distinct_staff_counts", index=False
        )
        codes_by_uoa.to_excel(writer, sheet_name="codes_by_uoa", index=False)
        codes_by_pathway.to_excel(writer, sheet_name="codes_by_pathway", index=False)
        df.to_excel(writer, sheet_name="master_rows", index=False)
        diagnostics.to_excel(
            writer, sheet_name="file_level_extraction", index=False
        )
        if not pivot_by_pathway_uoa.empty:
            pivot_by_pathway_uoa.to_excel(writer, sheet_name="pivot_pathway_uoa")

    return buf.getvalue()