"""
OPD Workup Scheduler — Real-Time Monitoring UI  (Production v3)
Variable-duration bookings: each appointment reserves exactly predict_workup() minutes.
Files needed (same folder):
  workup_model_age7_ar1.pkl, ehr_age7_ar1.xlsx,
  patient_lookup.csv (auto-built), bookings.json (auto-built)

Legacy bookings.json entries with a "slot" key are migrated automatically on first load.
"""

import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont, filedialog
import pandas as pd
import numpy as np
import pickle
import os
import json
import csv
import threading
import logging
from datetime import datetime

# ── Logging ───────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
logging.basicConfig(
    filename=os.path.join(BASE, "scheduler.log"),
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("opd")

# ── Paths ─────────────────────────────────────────────────────────────
PKL_PATH      = os.path.join(BASE, "workup_model_age7_ar1.pkl")
EHR_PATH      = os.path.join(BASE, "ehr_age7_ar1.xlsx")
LOOKUP_CACHE  = os.path.join(BASE, "patient_lookup.csv")
BOOKINGS_FILE = os.path.join(BASE, "bookings.json")

# ── Palette ───────────────────────────────────────────────────────────
BG        = "#0d1117"
CARD      = "#161b22"
CARD2     = "#1c2333"
BORDER    = "#30363d"
BORDER2   = "#21262d"
ACCENT    = "#58a6ff"
ACCENT2   = "#1f6feb"
GREEN     = "#3fb950"
GREEN_LT  = "#1a4a27"
GREEN2    = "#238636"
OT_COL    = "#6e7681"
TEXT      = "#e6edf3"
TEXT_DIM  = "#8b949e"
RED       = "#f85149"
TEAL      = "#39d353"
YELLOW    = "#d29922"
YELLOW_BG = "#3a2a00"
SIDEBAR   = "#010409"
PURPLE    = "#8957e5"

# ── Broad period definitions ──────────────────────────────────────────
SLOT_LABELS  = ["08-09","09-10","10-11","11-12","12-13:30",
                "13:30-14:30","14:30-15:30","15:30-17"]
SLOT_START_H = [8, 9, 10, 11, 12, 13.5, 14.5, 15.5]
SLOT_DUR_H   = [1, 1,  1,  1,  1.5,  1,    1,    1.5]

PERIOD_START_MIN = [int(h * 60) for h in SLOT_START_H]
PERIOD_END_MIN   = [int((h + d) * 60) for h, d in zip(SLOT_START_H, SLOT_DUR_H)]
PERIOD_DUR_MIN   = [e - s for s, e in zip(PERIOD_START_MIN, PERIOD_END_MIN)]

DAY_START_MIN = 8 * 60    # 08:00
DAY_END_MIN   = 17 * 60   # 17:00
DAY_SPAN_MIN  = DAY_END_MIN - DAY_START_MIN   # 540

SPECIALTIES = ["Retina","Glaucoma","Cornea","General","Pediatric and Low vision"]
DAYS_ORDER  = ["Mon","Tue","Wed","Thu","Fri","Sat"]
FLOW_TYPES  = ["Non-Dilated","Dilated","Procedure"]
AGE_CATS    = ["0-12","13-19","20-30","31-45","46-60","61-75","76-90"]
FLOW_ICONS  = {"Non-Dilated":"👁","Dilated":"👁‍🗨","Procedure":"⚡"}

# ── Schedule data ─────────────────────────────────────────────────────
SCHEDULE = {
    "Retina": {
        "Mon": [["OT","OT","OT","OT","OPD","LUNCH","NA","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","NA","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"]],
        "Tue": [["OT","OT","OT","OT","OPD","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","NA","LUNCH","NA","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","NA","LUNCH","OPD","OPD"]],
        "Wed": [["OT","OT","OT","OT","OPD","LUNCH","NA","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","NA","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","NA","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"]],
        "Thu": [["OPD","OPD","OPD","OPD","NA","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","NA","LUNCH","OPD","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"]],
        "Fri": [["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","NA","LUNCH","NA","OPD"]],
        "Sat": [["OPD","OPD","OPD","OPD","NA","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","NA","LUNCH","OPD","OPD"]],
    },
    "Glaucoma": {
        "Mon": [["OPD","OPD","OPD","OPD","NA","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"]],
        "Tue": [["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","NA","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"]],
        "Wed": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"]],
        "Thu": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","OPD","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","NA","LUNCH","OPD","OPD"]],
        "Fri": [["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"]],
        "Sat": [["OT","OT","OT","OT","NA","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"]],
    },
    "Cornea": {
        "Mon": [["OT","OT","OT","OT","OT","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"]],
        "Tue": [["OT","OT","OT","OT","OT","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"]],
        "Wed": [["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","NA","LUNCH","OPD","OPD"]],
        "Thu": [["OPD","OPD","OPD","OPD","NA","LUNCH","NA","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"]],
        "Fri": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"]],
        "Sat": [["OPD","OPD","OPD","NA","NA","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"]],
    },
    "General": {
        "Mon": [["OPD","OPD","OPD","NA","NA","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OPD","OPD","NA","NA","OPD","LUNCH","OPD","OPD"]],
        "Tue": [["OT","OT","OT","OT","OT","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","OPD"]],
        "Wed": [["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"]],
        "Thu": [["OT","OT","OT","OT","OT","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"]],
        "Fri": [["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"]],
        "Sat": [["OPD","OPD","OPD","OPD","OPD","LUNCH","NA","NA"],["OPD","OPD","OPD","NA","NA","LUNCH","NA","NA"],["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","NA","OPD"]],
    },
    "Pediatric and Low vision": {
        "Mon": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","NA","NA"]],
        "Tue": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","NA","OPD"]],
        "Wed": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"]],
        "Thu": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","NA","LUNCH","OPD","OPD"]],
        "Fri": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","NA"],["OT","OT","OT","OT","OPD","LUNCH","OPD","OPD"]],
        "Sat": [["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"],["OT","OT","OT","OT","OPD","LUNCH","NA","OPD"]],
    },
}

# ── Helpers ───────────────────────────────────────────────────────────
def fmt_range(mid):
    return f"{int(mid*0.85)}–{int(mid*1.15)} min"

def mins_to_hm(m):
    h, mn = divmod(int(m), 60)
    if h: return f"{h}h {mn}m" if mn else f"{h}h"
    return f"{mn} min"

def min_to_clock(m):
    h, mn = divmod(int(m), 60)
    return f"{h:02d}:{mn:02d}"

def clock_range(start_min, duration_min):
    return f"{min_to_clock(start_min)}–{min_to_clock(start_min + duration_min)}"

def _lerp_color(c1, c2, t):
    r1,g1,b1 = int(c1[1:3],16),int(c1[3:5],16),int(c1[5:7],16)
    r2,g2,b2 = int(c2[1:3],16),int(c2[3:5],16),int(c2[5:7],16)
    r=int(r1+(r2-r1)*t); g=int(g1+(g2-g1)*t); b=int(b1+(b2-b1)*t)
    return f"#{r:02x}{g:02x}{b:02x}"

# ── Period helpers ────────────────────────────────────────────────────
def get_broad_idx_from_minute(start_min):
    for i,(ps,pe) in enumerate(zip(PERIOD_START_MIN, PERIOD_END_MIN)):
        if ps <= start_min < pe:
            return i
    return -1

def get_opd_periods(doc_slots):
    return [(PERIOD_START_MIN[i], PERIOD_END_MIN[i])
            for i, s in enumerate(doc_slots) if s == "OPD"]

def total_opd_minutes(doc_slots):
    return sum(PERIOD_DUR_MIN[i] for i,s in enumerate(doc_slots) if s=="OPD")

def get_period_status_at(doc_slots, minute):
    idx = get_broad_idx_from_minute(minute)
    if 0 <= idx < len(doc_slots):
        return doc_slots[idx]
    return "NA"

# ── Booking migration ─────────────────────────────────────────────────
def _parse_old_slot_label(slot_label):
    try:
        start_str = slot_label.split("–")[0]
        h, m = map(int, start_str.split(":"))
        return h*60+m, 15
    except Exception:
        return 9*60, 15

def migrate_booking(b):
    if "start_minute" not in b and "slot" in b:
        sm, dur = _parse_old_slot_label(b["slot"])
        b["start_minute"]     = sm
        b["duration_minutes"] = dur
        b["end_minute"]       = sm + dur
    return b

# ── Booking store ─────────────────────────────────────────────────────
def load_bookings():
    if os.path.exists(BOOKINGS_FILE):
        with open(BOOKINGS_FILE) as f:
            raw = json.load(f)
        migrated = [migrate_booking(b) for b in raw]
        if any("slot" in b and "start_minute" in b for b in migrated):
            save_bookings(migrated)
        return migrated
    return []

def save_bookings(bookings):
    with open(BOOKINGS_FILE,"w") as f:
        json.dump(bookings, f, indent=2)

def is_already_booked(bookings, mrdno, day):
    for b in bookings:
        if str(b["mrdno"])==str(mrdno) and b["day"]==day:
            return b
    return None

def has_interval_overlap(bookings, doctor, day, start_min, duration_min, exclude_id=None):
    end_min = start_min + duration_min
    for b in bookings:
        if b.get("id")==exclude_id: continue
        if b["doctor"]==doctor and b["day"]==day:
            if start_min < b["end_minute"] and b["start_minute"] < end_min:
                return b
    return None

# ── Data loading ──────────────────────────────────────────────────────
def load_patient_lookup():
    if os.path.exists(LOOKUP_CACHE):
        return pd.read_csv(LOOKUP_CACHE, dtype={"MRDNO":str}).set_index("MRDNO")
    df = pd.read_excel(EHR_PATH)
    df["MRDNO"] = df["MRDNO"].astype(str)
    # Strip whitespace from categorical columns (EHR has trailing spaces)
    df["Patient Flow type"] = df["Patient Flow type"].astype(str).str.strip()
    df["Patient visit category"] = df["Patient visit category"].astype(str).str.strip()
    df["Specialty"] = df["Specialty"].astype(str).str.strip()

    def _to_min(x):
        if hasattr(x, "hour"): return x.hour*60 + x.minute
        return None

    df["workup_min"]  = df["TOTAL_WORKUP_TIME"].apply(_to_min)
    # Consultation duration: CONS_WORKUP_TIME is the actual doctor face time per visit
    df["consult_min"] = df["CONS_WORKUP_TIME"].apply(_to_min)
    df.loc[df["consult_min"] <= 0, "consult_min"] = None

    df_sorted = df.sort_values("VISITDATE", ascending=False)
    info = df_sorted.groupby("MRDNO").first()[[
        "Gender","Age category","Specialty","Patient visit category",
        "Patient Flow type","Consultant Name","Consultant Designation",
        "Day","Session","Arrival hour"]]
    vc         = df.groupby("MRDNO").size().rename("past_visits")
    avg_workup = df.groupby("MRDNO")["workup_min"].mean().rename("avg_workup")

    # Per-patient avg consultation time split by flow type
    # (a patient booked as Non-Dilated today should get their Non-Dilated avg, not a
    #  lifetime average inflated by past Procedure visits which have 74-min consults)
    flow_avgs = (df.groupby(["MRDNO","Patient Flow type"])["consult_min"]
                   .mean().unstack("Patient Flow type"))
    flow_avgs.columns = [f"avg_consult_{c.replace('-','').replace(' ','_').lower()}"
                         for c in flow_avgs.columns]

    lookup = info.join(vc).join(avg_workup).join(flow_avgs)
    lookup.reset_index().to_csv(LOOKUP_CACHE, index=False)
    return lookup

# Per-flow global median OPD slot durations (from EHR n=225,961 + capacity sanity)
# Non-Dilated : CONS_WORKUP_TIME median ~4 min — quick review, no dilation
# Dilated     : Split into TWO bookings:
#                 Phase 1 (initial check): 4 min — doctor checks, instils drops
#                 [gap of DILATION_WAIT_MIN — patient waits in lobby for drops to work]
#                 Phase 2 (fundus exam):   10 min — doctor does post-dilation fundus exam
#               Old EHR CONS_WORKUP_TIME ~38 min was the combined total; now split.
# Procedure   : CONS_WORKUP_TIME ~76 min but that covers the whole procedure suite stay
#               (drops + procedure + recovery).  The OPD consultation slot is only the
#               pre-procedure check; cap at 15 min so the doctor's OPD schedule fits
#               ~350 patients/day alongside OT blocks.
GLOBAL_CONSULT_MEDIAN = {
    "Non-Dilated": 4,
    "Dilated":     4,    # Phase 1 initial check only; Phase 2 exam = DILATION_EXAM_MIN
    "Procedure":   12,   # OPD pre-procedure check only; procedure itself is in OT
}
# Hard cap per flow — guards against outlier personal averages inflating slot sizes
CONSULT_SLOT_CAP = {
    "Non-Dilated": 30,
    "Dilated":     6,    # initial check capped at 6 min; EHR avg ~38 was full visit
    "Procedure":   15,   # even if a patient's avg is 150 min, OPD slot capped at 15
}

# Dilation two-slot parameters
DILATION_WAIT_MIN = 28   # biological minimum for drops to dilate pupils fully
DILATION_EXAM_MIN = 10   # post-dilation fundus exam duration

def get_consult_slot(patient_row, flow: str) -> int:
    """Return the OPD consultation slot (minutes) for this patient's flow type today."""
    flow = flow.strip()
    col  = f"avg_consult_{flow.replace('-','').replace(' ','_').lower()}"
    val  = patient_row.get(col)
    cap  = CONSULT_SLOT_CAP.get(flow, 60)
    if val is not None and pd.notna(val) and val > 0:
        return max(1, min(cap, round(float(val))))
    return GLOBAL_CONSULT_MEDIAN.get(flow, 10)

def load_model():
    class SafeUnpickler(pickle.Unpickler):
        def find_class(self, module, name):
            if "xgboost" in module:
                class Dummy: pass
                return Dummy
            return super().find_class(module, name)
    with open(PKL_PATH,"rb") as f:
        return SafeUnpickler(f).load()

def predict_workup(pkg, spec, flow, consultant, session, arrival, avg_workup):
    em = pkg["encoding_maps"]
    s  = em["spec_mean"].get(spec.strip(),  em["global_mean"])
    fl = em["flow_mean"].get(flow.strip(),  em["global_mean"])
    c  = em["consultant_mean"].get(consultant, em["global_mean"])
    bl = s*0.3 + fl*0.5 + c*0.2
    if session=="Afternoon": bl *= 0.95
    ah = em["arrival_map"].get(arrival, 10)
    if ah>=14: bl*=1.05
    elif ah<=9: bl*=0.95
    return round(bl*0.6 + avg_workup*0.4, 1)

# ── Plan B windows (from EHR analysis / DMAIC framework) ─────────────
# Follow-up (MRE/SRE) own the early windows; walk-ins (REG) get two
# dedicated slots: 9–11am and 2:30–5pm. REG must NOT intrude on the
# follow-up early/procedure blocks.
PLAN_B_WINDOWS = {
    # Follow-up — Procedure: must start 8–10am (120 min avg workup)
    ("SRE","Procedure"): [0, 1],
    ("MRE","Procedure"): [0, 1],
    # Follow-up — Dilated: 8am–12pm (72 min avg, needs dilation wait)
    ("SRE","Dilated"):   [0, 1, 2, 3],
    ("MRE","Dilated"):   [0, 1, 2, 3],
    # Follow-up — Non-Dilated: 10am–3pm flexible (39 min avg)
    ("SRE","Non-Dilated"): [2, 3, 4, 6, 7],
    ("MRE","Non-Dilated"): [2, 3, 4, 6, 7],
    # Walk-in (REG) — two dedicated windows only: 9–11am + 2:30–5pm
    ("REG","Non-Dilated"): [1, 2, 6, 7],
    ("REG","Dilated"):     [1, 2, 6, 7],
    ("REG","Procedure"):   [1, 2, 6, 7],
}

# Periods REG is allowed to search at all (hard constraint)
WALKIN_ALLOWED_PERIODS = {1, 2, 6, 7}

# Max fraction of an OPD period that REG can consume before follow-ups
# are fully protected — 37.5% mirrors the real patient mix
WALKIN_PERIOD_CAP = 0.375

ZONE_LABEL = {0:"Early",1:"Early",2:"Mid",3:"Mid",4:"Mid",5:"Lunch",6:"Afternoon",7:"Afternoon"}
ZONE_COLOR = {"Early":"#1a3a1f","Mid":"#1c2d3a","Lunch":"#3a2a00","Afternoon":"#1f1f3a"}
ZONE_FG    = {"Early":GREEN,"Mid":ACCENT,"Lunch":YELLOW,"Afternoon":PURPLE}

def _get_preferred_periods(visit_cat, flow):
    key = (visit_cat, flow.strip())
    if key in PLAN_B_WINDOWS: return PLAN_B_WINDOWS[key]
    if visit_cat == "REG": return list(WALKIN_ALLOWED_PERIODS)
    return list(range(8))

# ── Gap finder & scorer ───────────────────────────────────────────────
SCAN_STEP      = 5
OVERLAP_BUFFER = 3

def find_gaps(opd_periods, sorted_bookings, duration_needed):
    candidates = []
    for (ps, pe) in opd_periods:
        period_bk = [b for b in sorted_bookings
                     if b["start_minute"]<pe and b["end_minute"]>ps]
        t = ps
        while t + duration_needed <= pe:
            end_t = t + duration_needed
            overlaps = False
            for b in period_bk:
                if t < b["end_minute"] and b["start_minute"] < end_t:
                    overlaps = True
                    t = b["end_minute"]
                    rem = t % SCAN_STEP
                    if rem: t += SCAN_STEP - rem
                    break
            if overlaps: continue
            next_start = next((b["start_minute"] for b in period_bk
                               if b["start_minute"]>=end_t), None)
            if next_start is not None and end_t+OVERLAP_BUFFER > next_start:
                t += SCAN_STEP; continue
            candidates.append(t)
            t += SCAN_STEP
    return candidates

def find_dilation_pairs(slot1_periods, slot2_periods, sorted_bookings,
                        initial_dur, wait_min, exam_dur):
    """
    Find (check_start, exam_start) minute pairs for split dilated booking.
    slot1_periods: OPD periods eligible for the initial check (may be walk-in restricted).
    slot2_periods: all OPD periods eligible for the exam (patient already inside).
    Returns at most one pair per slot1 candidate (first valid exam slot wins).
    """
    pairs = []
    slot1_cands = find_gaps(slot1_periods, sorted_bookings, initial_dur)
    for t1 in slot1_cands:
        earliest_t2 = t1 + initial_dur + wait_min
        rem = earliest_t2 % SCAN_STEP
        if rem: earliest_t2 += SCAN_STEP - rem
        # Treat slot1 as occupied so exam doesn't overlap with it
        temp_bk = sorted(sorted_bookings + [
            {"start_minute": t1, "end_minute": t1 + initial_dur, "duration_minutes": initial_dur}
        ], key=lambda b: b["start_minute"])
        found = False
        for (ps, pe) in slot2_periods:
            if pe <= earliest_t2:
                continue
            t = max(ps, earliest_t2)
            rem2 = t % SCAN_STEP
            if rem2: t += SCAN_STEP - rem2
            period_bk = [b for b in temp_bk if b["start_minute"] < pe and b["end_minute"] > ps]
            while t + exam_dur <= pe:
                end_t = t + exam_dur
                overlap = False
                for b in period_bk:
                    if t < b["end_minute"] and b["start_minute"] < end_t:
                        overlap = True
                        t = b["end_minute"]
                        rem3 = t % SCAN_STEP
                        if rem3: t += SCAN_STEP - rem3
                        break
                if overlap: continue
                next_start = next((b["start_minute"] for b in period_bk
                                   if b["start_minute"] >= end_t), None)
                if next_start is not None and end_t + OVERLAP_BUFFER > next_start:
                    t += SCAN_STEP; continue
                pairs.append((t1, t))
                found = True
                break
            if found: break
    return pairs


def score_slot_v2(start_min, duration_min, doc_name, day, flow,
                  bookings, visit_cat, doc_slots):
    broad     = get_broad_idx_from_minute(start_min)
    preferred = _get_preferred_periods(visit_cat, flow)
    plan_b    = 0.0 if broad in preferred else 0.6
    cap       = total_opd_minutes(doc_slots)
    booked_m  = sum(b["duration_minutes"] for b in bookings
                    if b["day"]==day and b["doctor"]==doc_name)
    doc_load  = min(1.0, booked_m/max(cap,1))
    t_norm    = (start_min - 8*60)/(9*60)
    dil_pen   = 0.4 if flow.strip()=="Dilated" and t_norm>0.70 else 0.0
    lunch_pen = 0.3 if broad==5 else 0.0
    pd_dur    = PERIOD_DUR_MIN[broad] if 0<=broad<len(PERIOD_DUR_MIN) else 60
    fu_min    = sum(b["duration_minutes"] for b in bookings
                    if b["day"]==day
                    and get_broad_idx_from_minute(b["start_minute"])==broad
                    and b.get("visit_cat","") in ("SRE","MRE"))
    walkin    = 0.2 if fu_min > 0.8*pd_dur else 0.0
    score     = 0.40*plan_b + 0.25*doc_load + 0.10*dil_pen + 0.05*lunch_pen + 0.05*walkin
    return score, {"plan_b":plan_b,"doc_load":doc_load,"congestion":0.0,
                   "dil_pen":dil_pen,"lunch_pen":lunch_pen,
                   "preferred":broad in preferred,"zone":ZONE_LABEL.get(broad,""),
                   "doc_capacity_min":cap,"doc_booked_min":booked_m}

def rank_slots_v2(spec, day, flow, bookings, pred_duration, top_n=3, visit_cat="MRE"):
    sched = SCHEDULE.get(spec,{}).get(day,[])
    if not sched: return []
    duration = max(1, round(pred_duration))
    is_walkin = (visit_cat == "REG")

    # Per-period OPD capacity across ALL specialties (not just this one)
    # Must be cross-specialty to match how period_booked_reg is counted
    period_total_opd = [0]*8
    for _spec in SCHEDULE:
        for _ds in SCHEDULE[_spec].get(day, []):
            for i, s in enumerate(_ds):
                if s == "OPD": period_total_opd[i] += PERIOD_DUR_MIN[i]

    # How many minutes are already booked — split by type for capacity guard
    period_booked_reg = [0]*8   # REG minutes booked per period (all specialties)
    period_booked_all = [0]*8   # all minutes booked per period (all specialties)
    for b in bookings:
        if b["day"] == day:
            bi = get_broad_idx_from_minute(b["start_minute"])
            if 0 <= bi < 8:
                period_booked_all[bi] += b["duration_minutes"]
                if b.get("visit_cat") == "REG":
                    period_booked_reg[bi] += b["duration_minutes"]

    # Specialty prefix so "Retina Doc 1" ≠ "Glaucoma Doc 1" (different physical doctors)
    spec_prefix = spec.strip().split()[0][:3].upper()

    candidates = []
    for di, doc_slots in enumerate(sched):
        doc_name = f"{spec_prefix}-Doc{di+1}"
        all_opd  = get_opd_periods(doc_slots)
        if not all_opd: continue

        # Hard constraint: REG can only search their allowed periods
        if is_walkin:
            search_periods = [(ps, pe) for i, (ps, pe) in
                              enumerate(zip(PERIOD_START_MIN, PERIOD_END_MIN))
                              if i in WALKIN_ALLOWED_PERIODS
                              and i < len(doc_slots) and doc_slots[i] == "OPD"]
        else:
            # Follow-up: search all OPD periods
            # Only skip a period if REG has consumed MORE than WALKIN_PERIOD_CAP
            # (protects remaining follow-up capacity without blocking follow-ups entirely)
            search_periods = []
            for i, (ps, pe) in enumerate(zip(PERIOD_START_MIN, PERIOD_END_MIN)):
                if i >= len(doc_slots) or doc_slots[i] != "OPD": continue
                cap = period_total_opd[i]
                if cap > 0 and period_booked_reg[i] / cap > WALKIN_PERIOD_CAP:
                    continue   # REG already consumed their share of this period
                search_periods.append((ps, pe))

        if not search_periods: continue

        doc_bk = sorted([b for b in bookings if b["doctor"] == doc_name and b["day"] == day],
                        key=lambda b: b["start_minute"])

        if flow.strip() == "Dilated":
            # Two-slot booking: initial check + post-dilation fundus exam
            # slot1 restricted to search_periods; slot2 can use any OPD period
            pairs = find_dilation_pairs(search_periods, all_opd, doc_bk,
                                        duration, DILATION_WAIT_MIN, DILATION_EXAM_MIN)
            for (t1, t2) in pairs:
                broad = get_broad_idx_from_minute(t1)
                if is_walkin and period_total_opd[broad] > 0:
                    if (period_booked_reg[broad] + duration) / period_total_opd[broad] > WALKIN_PERIOD_CAP:
                        continue
                score, bd = score_slot_v2(t1, duration, doc_name, day, flow,
                                          bookings, visit_cat, doc_slots)
                cong = (min(1.0, period_booked_all[broad] / period_total_opd[broad])
                        if period_total_opd[broad] > 0 else 0.0)
                bd["congestion"] = cong
                bd["is_walkin"]  = is_walkin
                score += 0.15 * cong
                candidates.append({
                    "doc": doc_name, "start_minute": t1,
                    "duration_minutes": duration, "score": score,
                    "breakdown": bd, "clock_range": clock_range(t1, duration),
                    "exam_start_minute": t2,
                    "exam_duration_minutes": DILATION_EXAM_MIN,
                    "exam_clock_range": clock_range(t2, DILATION_EXAM_MIN),
                    "is_dilation_pair": True,
                })
        else:
            for sm in find_gaps(search_periods, doc_bk, duration):
                broad = get_broad_idx_from_minute(sm)

                # Capacity guard for REG: don't let a single walk-in push a period
                # past WALKIN_PERIOD_CAP of total OPD minutes
                if is_walkin and period_total_opd[broad] > 0:
                    if (period_booked_reg[broad] + duration) / period_total_opd[broad] > WALKIN_PERIOD_CAP:
                        continue

                score, bd = score_slot_v2(sm, duration, doc_name, day, flow,
                                          bookings, visit_cat, doc_slots)
                cong = (min(1.0, period_booked_all[broad] / period_total_opd[broad])
                        if period_total_opd[broad] > 0 else 0.0)
                bd["congestion"] = cong
                bd["is_walkin"]  = is_walkin
                score += 0.15 * cong
                candidates.append({"doc": doc_name, "start_minute": sm,
                                    "duration_minutes": duration, "score": score,
                                    "breakdown": bd, "clock_range": clock_range(sm, duration)})

    candidates.sort(key=lambda c: (c["score"], c["start_minute"], c["doc"]))
    return candidates[:top_n] if top_n else candidates

# ── Analytics ─────────────────────────────────────────────────────────
def compute_today_metrics(bookings):
    today = datetime.now().strftime("%a")
    if today not in DAYS_ORDER: today="Mon"
    today_bk = [b for b in bookings if b["day"]==today]
    total_min = sum(total_opd_minutes(ds)
                    for spec in SPECIALTIES
                    for ds in SCHEDULE.get(spec,{}).get(today,[]))
    booked_min = sum(b["duration_minutes"] for b in today_bk)
    # Count unique patients (not booking records — dilated has 2 records per patient)
    unique_patients = len({b["mrdno"] for b in today_bk})
    pts_per_hour = round(unique_patients / (DAY_SPAN_MIN / 60), 1) if DAY_SPAN_MIN else 0
    return {"day":today,"booked_ct":unique_patients,"booked_min":booked_min,
            "total_min":total_min,
            "pct":round(booked_min/total_min*100,1) if total_min else 0,
            "pts_per_hour": pts_per_hour,
            "by_spec":{sp:sum(b["duration_minutes"] for b in today_bk if b["specialty"]==sp)
                       for sp in SPECIALTIES},
            "by_flow":{ft:sum(1 for b in today_bk if b["flow"]==ft and b.get("dilation_phase")!="exam")
                       for ft in FLOW_TYPES}}

def compute_spec_day_pct(bookings, day, spec):
    sched = SCHEDULE.get(spec,{}).get(day,[])
    total = sum(total_opd_minutes(ds) for ds in sched)
    booked = sum(b["duration_minutes"] for b in bookings
                 if b["day"]==day and b["specialty"]==spec)
    return round(booked/total*100,1) if total else 0.0

# ── Toast ─────────────────────────────────────────────────────────────
class ToastManager:
    def __init__(self, root):
        self.root = root; self._queue = []; self._showing = False

    def show(self, message, kind="info", duration=3500):
        colors = {"info":(CARD2,ACCENT),"success":(GREEN_LT,GREEN),
                  "error":("#3a1a1a",RED),"warn":(YELLOW_BG,YELLOW)}
        bg, fg = colors.get(kind,(CARD2,TEXT))
        self._queue.append((message, kind, bg, fg, duration))
        if not self._showing: self._next()

    def _next(self):
        if not self._queue: self._showing=False; return
        self._showing = True
        message, kind, bg, fg, duration = self._queue.pop(0)
        toast = tk.Frame(self.root, bg=bg, padx=16, pady=10,
                         highlightbackground=fg, highlightthickness=1)
        sw = self.root.winfo_screenwidth(); sh = self.root.winfo_screenheight()
        toast.place(x=sw-400, y=sh-130); toast.lift()
        icon = {"info":"ℹ","success":"✓","error":"✗","warn":"⚠"}.get(kind,"ℹ")
        tk.Label(toast, text=f"{icon}  {message}", bg=bg, fg=fg,
                 font=("Segoe UI",10), wraplength=340, justify="left").pack(anchor="w")
        def dismiss():
            try: toast.place_forget(); toast.destroy()
            except Exception: pass
            self.root.after(200, self._next)
        self.root.after(duration, dismiss)
        toast.bind("<Button-1>", lambda e: dismiss())


# ═══════════════════════════════════════════════════════════════════════
# Main App
# ═══════════════════════════════════════════════════════════════════════
class OPDSchedulerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OPD Workup Scheduler  ·  Real-Time Monitoring")
        self.geometry("1440x920")
        self.minsize(1100, 700)
        self.configure(bg=BG)
        self.resizable(True,True)

        self.patient_data = None
        self.pkg          = None
        self.lookup       = None
        self.bookings     = load_bookings()
        self.selected_day = tk.StringVar(master=self, value=self._today_day())
        self.active_view  = tk.StringVar(master=self, value="scheduler")
        self._search_var  = tk.StringVar(master=self)
        self._last_save   = datetime.now()

        self._loading_dots=0; self._pulse_val=0
        self._loading_job=None; self._dot_job=None; self._clock_job=None

        self.toast = ToastManager(self)
        self._build_fonts()
        self._build_ui()
        self._bind_shortcuts()
        self._load_data_async()
        self._tick_clock()
        self.after(30_000, self._auto_refresh)

    def _today_day(self):
        d = datetime.now().strftime("%a")
        return d if d in DAYS_ORDER else "Mon"

    def _build_fonts(self):
        self.fnt_title = tkfont.Font(family="Segoe UI", size=18, weight="bold")
        self.fnt_head  = tkfont.Font(family="Segoe UI", size=12, weight="bold")
        self.fnt_body  = tkfont.Font(family="Segoe UI", size=11)
        self.fnt_small = tkfont.Font(family="Segoe UI", size=9)
        self.fnt_big   = tkfont.Font(family="Segoe UI", size=26, weight="bold")
        self.fnt_label = tkfont.Font(family="Segoe UI", size=10)
        self.fnt_nav   = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        self.fnt_mono  = tkfont.Font(family="Consolas",  size=10)
        self.fnt_tag   = tkfont.Font(family="Segoe UI", size=8,  weight="bold")
        self.fnt_num   = tkfont.Font(family="Segoe UI", size=22, weight="bold")
        self.fnt_clock = tkfont.Font(family="Consolas",  size=11, weight="bold")

    def _bind_shortcuts(self):
        self.bind_all("<F5>",        lambda e: self._refresh_current_view())
        self.bind_all("<Control-f>", lambda e: self._focus_search())
        self.bind_all("<Control-e>", lambda e: self._export_csv())
        self.bind_all("<Escape>",    lambda e: self._search_var.set(""))

    def _tick_clock(self):
        now = datetime.now()
        try:
            self.lbl_clock.config(text=now.strftime("%H:%M:%S"))
            self.lbl_date.config(text=now.strftime("%a  %d %b %Y"))
        except Exception: pass
        self._clock_job = self.after(1000, self._tick_clock)

    def _auto_refresh(self):
        self._refresh_current_view()
        self.after(30_000, self._auto_refresh)

    def _refresh_current_view(self):
        v = self.active_view.get()
        if v=="dashboard":    self._refresh_dashboard()
        elif v=="appointments": self._refresh_appointments()
        elif v=="scheduler" and self.patient_data: self._render_patient()

    # ── Animation ─────────────────────────────────────────────────────
    def _animate_hover_in(self, w, bf, bt, ff, ft, step=0, steps=8):
        t=step/steps; w.config(bg=_lerp_color(bf,bt,t),fg=_lerp_color(ff,ft,t))
        if step<steps: w._hover_job=self.after(16,lambda:self._animate_hover_in(w,bf,bt,ff,ft,step+1,steps))

    def _animate_hover_out(self, w, bf, bt, ff, ft, step=0, steps=8):
        t=step/steps; w.config(bg=_lerp_color(bf,bt,t),fg=_lerp_color(ff,ft,t))
        if step<steps: w._hover_job=self.after(16,lambda:self._animate_hover_out(w,bf,bt,ff,ft,step+1,steps))

    def _cancel_hover(self, w):
        if hasattr(w,"_hover_job") and w._hover_job:
            self.after_cancel(w._hover_job); w._hover_job=None

    def _start_loading_pulse(self):
        cols=[TEXT_DIM,YELLOW,"#e6a817",YELLOW,TEXT_DIM]
        self._pulse_val=(self._pulse_val+1)%len(cols)
        try: self.lbl_status.config(fg=cols[self._pulse_val])
        except Exception: return
        self._loading_job=self.after(400,self._start_loading_pulse)

    def _stop_loading_pulse(self):
        if self._loading_job: self.after_cancel(self._loading_job); self._loading_job=None

    def _start_dot_animation(self, label, base="Loading"):
        self._loading_dots=(self._loading_dots+1)%4
        try: label.config(text=base+"."*self._loading_dots)
        except Exception: return
        self._dot_job=self.after(500,lambda:self._start_dot_animation(label,base))

    def _stop_dot_animation(self):
        if self._dot_job: self.after_cancel(self._dot_job); self._dot_job=None

    def _flash_ready(self):
        cols=[GREEN,"#5fff78",GREEN,"#2ea043",GREEN]
        def step(i=0):
            if i<len(cols):
                try: self.lbl_status.config(fg=cols[i])
                except Exception: return
                self.after(120,lambda:step(i+1))
        step()

    # ── UI skeleton ───────────────────────────────────────────────────
    def _build_ui(self):
        self.sidebar = tk.Frame(self, bg=SIDEBAR, width=210)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        logo_wrap = tk.Frame(self.sidebar, bg=SIDEBAR)
        logo_wrap.pack(fill="x")
        tk.Frame(logo_wrap, bg=ACCENT, width=3).pack(side="left", fill="y", pady=18)
        logo_inner = tk.Frame(logo_wrap, bg=SIDEBAR)
        logo_inner.pack(side="left", fill="both", expand=True, padx=(10,0))
        tk.Label(logo_inner, text="🏥  OPD", font=self.fnt_title,
                 bg=SIDEBAR, fg=TEXT).pack(pady=(18,2), anchor="w")
        tk.Label(logo_inner, text="Workup Scheduler  v3.0",
                 font=self.fnt_small, bg=SIDEBAR, fg=TEXT_DIM).pack(anchor="w", pady=(0,14))

        tk.Frame(self.sidebar, bg=BORDER2, height=1).pack(fill="x", padx=12)
        clk = tk.Frame(self.sidebar, bg=SIDEBAR)
        clk.pack(fill="x", padx=16, pady=(10,4))
        self.lbl_clock = tk.Label(clk, text="--:--:--", font=self.fnt_clock,
                                   bg=SIDEBAR, fg=ACCENT)
        self.lbl_clock.pack(anchor="w")
        self.lbl_date = tk.Label(clk, text="", font=self.fnt_small,
                                  bg=SIDEBAR, fg=TEXT_DIM)
        self.lbl_date.pack(anchor="w")

        tk.Frame(self.sidebar, bg=BORDER2, height=1).pack(fill="x", padx=12, pady=(10,0))
        tk.Label(self.sidebar, text="NAVIGATION", font=self.fnt_tag,
                 bg=SIDEBAR, fg="#444c56").pack(anchor="w", padx=20, pady=(10,4))
        self._nav_btn("📊   Dashboard",    "dashboard")
        self._nav_btn("📋   Scheduler",    "scheduler")
        self._nav_btn("📅   Appointments", "appointments")

        tk.Frame(self.sidebar, bg=BORDER2, height=1).pack(fill="x", padx=12, pady=(16,4))
        tk.Label(self.sidebar, text="SHORTCUTS", font=self.fnt_tag,
                 bg=SIDEBAR, fg="#444c56").pack(anchor="w", padx=20, pady=(4,4))
        for key, desc in [("F5","Refresh"),("Ctrl+F","Search"),("Ctrl+E","Export CSV"),("Esc","Clear search")]:
            r = tk.Frame(self.sidebar, bg=SIDEBAR)
            r.pack(fill="x", padx=16, pady=1)
            tk.Label(r, text=key, font=self.fnt_small, bg=CARD2, fg=ACCENT, padx=4, pady=1).pack(side="left")
            tk.Label(r, text=f"  {desc}", font=self.fnt_small, bg=SIDEBAR, fg=TEXT_DIM).pack(side="left")

        tk.Frame(self.sidebar, bg=BORDER2, height=1).pack(side="bottom", fill="x", padx=12, pady=(0,4))
        sw = tk.Frame(self.sidebar, bg=SIDEBAR)
        sw.pack(side="bottom", fill="x", padx=14, pady=(0,12))
        df = tk.Frame(sw, bg=SIDEBAR)
        df.pack(anchor="w")
        self.lbl_status_dot = tk.Label(df, text="●", font=self.fnt_small, bg=SIDEBAR, fg=YELLOW)
        self.lbl_status_dot.pack(side="left")
        self.lbl_status = tk.Label(df, text="Loading", font=self.fnt_small, bg=SIDEBAR, fg=YELLOW)
        self.lbl_status.pack(side="left", padx=(4,0))
        self._start_dot_animation(self.lbl_status)
        self._start_loading_pulse()

        self.main = tk.Frame(self, bg=BG)
        self.main.pack(side="left", fill="both", expand=True)
        self._build_statusbar()

        self.page_dashboard    = tk.Frame(self.main, bg=BG)
        self.page_scheduler    = tk.Frame(self.main, bg=BG)
        self.page_appointments = tk.Frame(self.main, bg=BG)
        self._build_dashboard_page()
        self._build_scheduler_page()
        self._build_appointments_page()
        self._show_view("scheduler")

    def _build_statusbar(self):
        bar = tk.Frame(self.main, bg=CARD2, height=26)
        bar.pack(side="bottom", fill="x")
        bar.pack_propagate(False)
        self.lbl_sb = tk.Label(bar, text="", font=self.fnt_small, bg=CARD2, fg=TEXT_DIM)
        self.lbl_sb.pack(side="left", padx=14)
        self.lbl_sb_save = tk.Label(bar, text="", font=self.fnt_small, bg=CARD2, fg=TEXT_DIM)
        self.lbl_sb_save.pack(side="right", padx=14)
        tk.Label(bar, text="F5 Refresh  ·  Ctrl+F Search  ·  Ctrl+E Export",
                 font=self.fnt_small, bg=CARD2, fg="#444c56").pack(side="right", padx=12)
        self._tick_statusbar()

    def _tick_statusbar(self):
        total = len({b["mrdno"] for b in self.bookings})
        today = self._today_day()
        ct = len({b["mrdno"] for b in self.bookings if b["day"]==today})
        self.lbl_sb.config(text=f"Total patients: {total}   ·   Today ({today}): {ct}")
        self.lbl_sb_save.config(text=f"Last save: {self._last_save.strftime('%H:%M:%S')}")
        self.after(5000, self._tick_statusbar)

    def _nav_btn(self, label, view):
        btn = tk.Label(self.sidebar, text=label, font=self.fnt_nav,
                       bg=SIDEBAR, fg=TEXT_DIM, anchor="w", padx=20, pady=11, cursor="hand2")
        btn.pack(fill="x")
        btn._view = view; btn._hover_job = None
        def on_enter(e,b=btn):
            self._cancel_hover(b)
            if self.active_view.get()!=b._view: self._animate_hover_in(b,SIDEBAR,CARD2,TEXT_DIM,TEXT)
        def on_leave(e,b=btn):
            self._cancel_hover(b)
            if self.active_view.get()!=b._view: self._animate_hover_out(b,CARD2,SIDEBAR,TEXT,TEXT_DIM)
        btn.bind("<Enter>",    on_enter)
        btn.bind("<Leave>",    on_leave)
        btn.bind("<Button-1>", lambda e,v=view: self._show_view(v))
        if not hasattr(self,"_nav_btns"): self._nav_btns=[]
        self._nav_btns.append(btn)

    def _show_view(self, view):
        self.active_view.set(view)
        for p in [self.page_dashboard,self.page_scheduler,self.page_appointments]:
            p.pack_forget()
        if view=="dashboard":
            self.page_dashboard.pack(fill="both",expand=True); self._refresh_dashboard()
        elif view=="scheduler":
            self.page_scheduler.pack(fill="both",expand=True)
        else:
            self.page_appointments.pack(fill="both",expand=True); self._refresh_appointments()
        for b in getattr(self,"_nav_btns",[]):
            b.config(bg=CARD2 if b._view==view else SIDEBAR,
                     fg=ACCENT if b._view==view else TEXT_DIM)

    # ═══════════════════════════════════════════════════════════════════
    # Dashboard
    # ═══════════════════════════════════════════════════════════════════
    def _build_dashboard_page(self):
        p = self.page_dashboard
        hdr = tk.Frame(p, bg=CARD); hdr.pack(fill="x")
        hi  = tk.Frame(hdr, bg=CARD, pady=12); hi.pack(fill="x", padx=16)
        tk.Label(hi, text="📊  Real-Time Monitoring Dashboard",
                 font=self.fnt_title, bg=CARD, fg=TEXT).pack(side="left")
        self.lbl_dash_ts = tk.Label(hi, text="", font=self.fnt_small, bg=CARD, fg=TEXT_DIM)
        self.lbl_dash_ts.pack(side="right")
        tk.Frame(p, bg=ACCENT2, height=2).pack(fill="x")
        canv = tk.Canvas(p, bg=BG, highlightthickness=0)
        vsb  = tk.Scrollbar(p, orient="vertical", command=canv.yview)
        canv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y"); canv.pack(fill="both", expand=True)
        self._dash_body = tk.Frame(canv, bg=BG)
        self._dash_win  = canv.create_window((0,0), window=self._dash_body, anchor="nw")
        self._dash_body.bind("<Configure>", lambda e: canv.configure(scrollregion=canv.bbox("all")))
        canv.bind("<Configure>", lambda e: canv.itemconfig(self._dash_win, width=e.width))
        canv.bind_all("<MouseWheel>", lambda e: canv.yview_scroll(
            int(-1*(e.delta/120)),"units") if self.active_view.get()=="dashboard" else None)

    def _refresh_dashboard(self):
        for w in self._dash_body.winfo_children(): w.destroy()
        m   = compute_today_metrics(self.bookings)
        day = m["day"]
        self.lbl_dash_ts.config(text=f"Updated  {datetime.now().strftime('%H:%M:%S')}  ·  30s auto-refresh")

        # KPI row
        kpi_row = tk.Frame(self._dash_body, bg=BG)
        kpi_row.pack(fill="x", padx=16, pady=(16,8))
        def kpi(parent, title, val, sub, col):
            o = tk.Frame(parent, bg=col, padx=1, pady=1)
            o.pack(side="left", fill="both", expand=True, padx=6)
            i = tk.Frame(o, bg=CARD, padx=16, pady=14); i.pack(fill="both")
            tk.Label(i, text=title, font=self.fnt_small, bg=CARD, fg=TEXT_DIM).pack(anchor="w")
            tk.Label(i, text=str(val), font=self.fnt_num,  bg=CARD, fg=col).pack(anchor="w")
            tk.Label(i, text=sub,      font=self.fnt_small, bg=CARD, fg=TEXT_DIM).pack(anchor="w")
        all_unique = len({b["mrdno"] for b in self.bookings})
        kpi(kpi_row,"PATIENTS TODAY",  m["booked_ct"],  f"Day: {day}",                 ACCENT)
        kpi(kpi_row,"PATIENTS/HOUR",   f"{m['pts_per_hour']}", "across 9h OPD day",   GREEN)
        kpi(kpi_row,"UTILIZATION",     f"{m['pct']}%",  "booked min / OPD capacity",  YELLOW)
        kpi(kpi_row,"ALL PATIENTS",    all_unique,      "unique across all days",      PURPLE)

        # Specialty bars for today
        self._section(self._dash_body, f"SPECIALTY UTILIZATION — {day.upper()}", ACCENT2)
        for spec in SPECIALTIES:
            pct = compute_spec_day_pct(self.bookings, day, spec)
            bc  = GREEN if pct<60 else YELLOW if pct<85 else RED
            row = tk.Frame(self._dash_body, bg=CARD, pady=8, padx=14,
                           highlightbackground=BORDER2, highlightthickness=1)
            row.pack(fill="x", padx=16, pady=3)
            nc = tk.Frame(row, bg=CARD, width=200); nc.pack(side="left"); nc.pack_propagate(False)
            tk.Label(nc, text=spec, font=self.fnt_label, bg=CARD, fg=TEXT, anchor="w").pack(fill="x")
            bc2 = tk.Frame(row, bg=CARD); bc2.pack(side="left", fill="both", expand=True)
            track = tk.Frame(bc2, bg=BORDER2, height=12); track.pack(fill="x", pady=(4,2))
            track.update_idletasks()
            tk.Frame(track, bg=bc, height=12).place(x=0,y=0,relwidth=min(1.0,pct/100),height=12)
            tk.Label(bc2, text=f"{pct}% utilized", font=self.fnt_small, bg=CARD, fg=TEXT_DIM).pack(anchor="w")
            tk.Label(row, text=f"{pct:.0f}%", font=self.fnt_head, bg=CARD, fg=bc, width=6).pack(side="right")

        # Weekly heatmap
        self._section(self._dash_body, "WEEKLY UTILIZATION HEATMAP", GREEN2)
        heat = tk.Frame(self._dash_body, bg=CARD, padx=2, pady=2,
                        highlightbackground=BORDER2, highlightthickness=1)
        heat.pack(fill="x", padx=16, pady=(0,8))
        hr = tk.Frame(heat, bg=CARD2); hr.pack(fill="x")
        tk.Label(hr, text="", bg=CARD2, width=24).pack(side="left")
        for d in DAYS_ORDER:
            tk.Label(hr, text=d, font=self.fnt_tag, bg=CARD2, fg=TEXT_DIM,
                     width=9, anchor="center", pady=6).pack(side="left")
        for spec in SPECIALTIES:
            row = tk.Frame(heat, bg=CARD); row.pack(fill="x")
            tk.Label(row, text=spec[:22], font=self.fnt_small, bg=CARD, fg=TEXT_DIM,
                     width=24, anchor="w", padx=8, pady=6).pack(side="left")
            for d in DAYS_ORDER:
                pct = compute_spec_day_pct(self.bookings, d, spec)
                cbg = _lerp_color("#1a3a1f","#f85149",pct/100) if pct>0 else CARD2
                tk.Label(row, text=f"{pct:.0f}%", font=self.fnt_small,
                         bg=cbg, fg=TEXT if pct>40 else TEXT_DIM,
                         width=9, anchor="center", pady=6).pack(side="left")
            tk.Frame(heat, bg=BORDER2, height=1).pack(fill="x")

        # Flow breakdown
        self._section(self._dash_body, f"FLOW TYPE BREAKDOWN — {day.upper()}", YELLOW)
        fr = tk.Frame(self._dash_body, bg=BG); fr.pack(fill="x", padx=16, pady=(0,8))
        fc = {"Non-Dilated":TEAL,"Dilated":YELLOW,"Procedure":RED}
        total_ct = max(1, m["booked_ct"])
        for ft in FLOW_TYPES:
            cnt = m["by_flow"].get(ft,0)
            pct = round(cnt/total_ct*100,1)
            col = fc.get(ft,TEXT)
            o=tk.Frame(fr,bg=col,padx=1,pady=1); o.pack(side="left",fill="both",expand=True,padx=6)
            i=tk.Frame(o,bg=CARD,padx=14,pady=12); i.pack(fill="both")
            tk.Label(i,text=FLOW_ICONS.get(ft,""),font=("Segoe UI",14),bg=CARD,fg=col).pack(anchor="w")
            tk.Label(i,text=ft,font=self.fnt_small,bg=CARD,fg=TEXT_DIM).pack(anchor="w")
            tk.Label(i,text=str(cnt),font=self.fnt_num,bg=CARD,fg=col).pack(anchor="w")
            tk.Label(i,text=f"{pct}% of today",font=self.fnt_small,bg=CARD,fg=TEXT_DIM).pack(anchor="w")

        # Recent bookings
        self._section(self._dash_body, "RECENT BOOKINGS", PURPLE)
        feed = tk.Frame(self._dash_body, bg=CARD, highlightbackground=BORDER2, highlightthickness=1)
        feed.pack(fill="x", padx=16, pady=(0,16))
        recent = sorted(self.bookings, key=lambda b: b.get("booked_at",""), reverse=True)[:12]
        if not recent:
            tk.Label(feed, text="No bookings yet.", font=self.fnt_body,
                     bg=CARD, fg=TEXT_DIM, pady=20).pack()
        else:
            hrow = tk.Frame(feed, bg=CARD2); hrow.pack(fill="x")
            for ht, w in [("MRDNO",10),("SPECIALTY",18),("FLOW",14),
                          ("DOCTOR",8),("DAY",6),("TIME WINDOW",18),("BOOKED AT",18)]:
                tk.Label(hrow,text=ht,font=self.fnt_tag,bg=CARD2,fg="#444c56",
                         width=w,anchor="w",pady=7,padx=8).pack(side="left")
            for i,b in enumerate(recent):
                rbg=CARD if i%2==0 else "#0f1419"
                row=tk.Frame(feed,bg=rbg); row.pack(fill="x")
                tw=clock_range(b.get("start_minute",0),b.get("duration_minutes",15))
                fc2={"Non-Dilated":TEAL,"Dilated":YELLOW,"Procedure":RED}.get(b.get("flow",""),TEXT)
                for val,w,fg in [(b.get("mrdno",""),10,TEXT),(b.get("specialty",""),18,TEXT_DIM),
                                  (b.get("flow",""),14,fc2),(b.get("doctor",""),8,TEXT_DIM),
                                  (b.get("day",""),6,ACCENT),(tw,18,TEXT),
                                  (b.get("booked_at","")[:16],18,TEXT_DIM)]:
                    tk.Label(row,text=str(val),font=self.fnt_small,bg=rbg,fg=fg,
                             width=w,anchor="w",pady=7,padx=8).pack(side="left")
                tk.Frame(feed,bg=BORDER2,height=1).pack(fill="x")

    def _section(self, parent, text, col):
        f = tk.Frame(parent, bg=BG); f.pack(fill="x", padx=16, pady=(14,4))
        tk.Frame(f, bg=col, width=3).pack(side="left", fill="y")
        tk.Label(f, text=f"  {text}", font=self.fnt_tag, bg=BG, fg=TEXT_DIM).pack(side="left")

    # ═══════════════════════════════════════════════════════════════════
    # Scheduler page
    # ═══════════════════════════════════════════════════════════════════
    def _build_scheduler_page(self):
        p = self.page_scheduler
        top = tk.Frame(p, bg=CARD); top.pack(fill="x")
        inner = tk.Frame(top, bg=CARD, pady=10); inner.pack(fill="x", padx=16)

        mrd_grp = tk.Frame(inner, bg=CARD2, padx=2, pady=2,
                           highlightbackground=BORDER, highlightthickness=1)
        mrd_grp.pack(side="left")
        tk.Label(mrd_grp, text=" MRDNO ", font=self.fnt_small,
                 bg=ACCENT2, fg=TEXT, padx=6).pack(side="left")
        self.entry_mrdno = tk.Entry(mrd_grp, font=self.fnt_head, width=14,
                                     bg=CARD2, fg=TEXT, insertbackground=ACCENT,
                                     relief="flat", bd=4)
        self.entry_mrdno.pack(side="left", padx=(4,6))
        self.entry_mrdno.bind("<Return>", lambda e: self._lookup())
        self.entry_mrdno.bind("<FocusIn>",  lambda e: mrd_grp.config(highlightbackground=ACCENT))
        self.entry_mrdno.bind("<FocusOut>", lambda e: mrd_grp.config(highlightbackground=BORDER))

        day_grp = tk.Frame(inner, bg=CARD2, padx=2, pady=2,
                           highlightbackground=BORDER, highlightthickness=1)
        day_grp.pack(side="left", padx=(10,0))
        tk.Label(day_grp, text=" Day ", font=self.fnt_small,
                 bg=GREEN2, fg=TEXT, padx=6).pack(side="left")
        self.day_cb = ttk.Combobox(day_grp, textvariable=self.selected_day,
                                    values=DAYS_ORDER, width=6, font=self.fnt_body, state="readonly")
        self.day_cb.pack(side="left", padx=(4,6))
        self.day_cb.bind("<<ComboboxSelected>>",
                          lambda e: self._render_patient() if self.patient_data else None)

        def _btn(text, bg, fg, hbg, hfg, cmd, hl=None):
            b = tk.Label(inner, text=text, font=self.fnt_body, bg=bg, fg=fg,
                         cursor="hand2", padx=6, pady=5,
                         highlightbackground=hl or bg, highlightthickness=1)
            b.pack(side="left", padx=(8,0))
            b.bind("<Button-1>", lambda e: cmd())
            b.bind("<Enter>",    lambda e: b.config(bg=hbg, fg=hfg))
            b.bind("<Leave>",    lambda e: b.config(bg=bg,  fg=fg))
        _btn("  Look up  ", ACCENT2, TEXT, ACCENT, BG, self._lookup, ACCENT)
        _btn("  Manual Entry  ", CARD, TEXT_DIM, CARD2, TEXT, self._open_manual, BORDER)
        _btn("  ⚡ Recommend  ", YELLOW_BG, YELLOW, YELLOW, BG, self._recommend_slot, YELLOW)
        _btn("  🗑 Clear  ", CARD, TEXT_DIM, "#3a1a1a", RED, self._clear_patient, BORDER)

        tk.Frame(p, bg=ACCENT2, height=2).pack(fill="x")

        self.canvas = tk.Canvas(p, bg=BG, highlightthickness=0)
        vsb = tk.Scrollbar(p, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y"); self.canvas.pack(fill="both", expand=True)
        self.content = tk.Frame(self.canvas, bg=BG)
        self.content_win = self.canvas.create_window((0,0), window=self.content, anchor="nw")
        self.content.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            self.content_win, width=e.width))
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(
            int(-1*(e.delta/120)),"units") if self.active_view.get()=="scheduler" else None)
        self._show_placeholder()

    def _show_placeholder(self):
        ph = tk.Frame(self.content, bg=BG)
        ph.pack(expand=True, fill="both", pady=100)
        tk.Label(ph, text="🔍", font=("Segoe UI",32), bg=BG, fg=BORDER).pack()
        tk.Label(ph, text="Enter a MRDNO and click  Look up",
                 font=self.fnt_body, bg=BG, fg=TEXT_DIM).pack(pady=(8,0))
        tk.Label(ph, text="or use Manual Entry to enter patient details directly",
                 font=self.fnt_small, bg=BG, fg="#444c56").pack(pady=(4,0))

    def _clear_patient(self):
        self.patient_data = None
        self.entry_mrdno.delete(0, tk.END)
        for w in self.content.winfo_children(): w.destroy()
        self._show_placeholder()

    # ═══════════════════════════════════════════════════════════════════
    # Appointments page
    # ═══════════════════════════════════════════════════════════════════
    def _build_appointments_page(self):
        p = self.page_appointments
        hdr = tk.Frame(p, bg=CARD); hdr.pack(fill="x")
        hi  = tk.Frame(hdr, bg=CARD, pady=12); hi.pack(fill="x", padx=16)
        tk.Label(hi, text="📅  All Appointments", font=self.fnt_title, bg=CARD, fg=TEXT).pack(side="left")
        tk.Frame(p, bg=ACCENT2, height=2).pack(fill="x")

        fbar = tk.Frame(p, bg=BG, pady=10); fbar.pack(fill="x", padx=16)
        sw   = tk.Frame(fbar, bg=CARD2, padx=2, pady=2,
                        highlightbackground=BORDER, highlightthickness=1)
        sw.pack(side="left")
        tk.Label(sw, text=" 🔍 ", font=self.fnt_small, bg=CARD2, fg=TEXT_DIM).pack(side="left")
        self._appt_search = tk.Entry(sw, textvariable=self._search_var, font=self.fnt_label,
                                      width=16, bg=CARD2, fg=TEXT, insertbackground=ACCENT,
                                      relief="flat", bd=4)
        self._appt_search.pack(side="left", padx=(0,6))
        self._appt_search.bind("<KeyRelease>", lambda e: self._refresh_appointments())
        self._appt_search.bind("<FocusIn>",  lambda e: sw.config(highlightbackground=ACCENT))
        self._appt_search.bind("<FocusOut>", lambda e: sw.config(highlightbackground=BORDER))

        tk.Label(fbar, text=" Day:", font=self.fnt_small, bg=BG, fg=TEXT_DIM).pack(side="left", padx=(10,0))
        self.appt_day_var = tk.StringVar(master=self, value="All")
        ttk.Combobox(fbar, textvariable=self.appt_day_var, values=["All"]+DAYS_ORDER,
                     width=7, font=self.fnt_label, state="readonly").pack(side="left", padx=(4,0))
        self.appt_day_var.trace_add("write", lambda *_: self._refresh_appointments())

        tk.Label(fbar, text=" Spec:", font=self.fnt_small, bg=BG, fg=TEXT_DIM).pack(side="left", padx=(10,0))
        self.appt_spec_var = tk.StringVar(master=self, value="All")
        ttk.Combobox(fbar, textvariable=self.appt_spec_var, values=["All"]+SPECIALTIES,
                     width=18, font=self.fnt_label, state="readonly").pack(side="left", padx=(4,0))
        self.appt_spec_var.trace_add("write", lambda *_: self._refresh_appointments())

        def _sm(text, bg, fg, hbg, hfg, cmd):
            b = tk.Label(fbar, text=text, font=self.fnt_small, bg=bg, fg=fg,
                         cursor="hand2", padx=6, pady=4, highlightbackground=BORDER, highlightthickness=1)
            b.pack(side="left", padx=(8,0))
            b.bind("<Button-1>", lambda e: cmd())
            b.bind("<Enter>",    lambda e: b.config(bg=hbg, fg=hfg))
            b.bind("<Leave>",    lambda e: b.config(bg=bg,  fg=fg))
        _sm("  Clear  ", CARD, TEXT_DIM, CARD2, TEXT, self._clear_appt_filters)
        _sm("  ⬇ Export CSV  ", GREEN_LT, GREEN, "#2a5a2a", GREEN, self._export_csv)

        self.appt_count_lbl = tk.Label(fbar, text="", font=self.fnt_small, bg=BG, fg=TEXT_DIM)
        self.appt_count_lbl.pack(side="right")
        tk.Frame(p, bg=BORDER2, height=1).pack(fill="x")

        cols   = ["MRDNO","Patient","Specialty","Flow","Doctor","Day","Time Window","Duration","Booked At",""]
        widths = [10,14,16,14,8,6,16,8,16,8]
        hrow   = tk.Frame(p, bg=CARD2); hrow.pack(fill="x")
        for c,w in zip(cols,widths):
            tk.Label(hrow, text=c, font=self.fnt_tag, bg=CARD2, fg="#444c56",
                     width=w, anchor="w", pady=8, padx=8).pack(side="left")
        tk.Frame(p, bg=ACCENT2, height=1).pack(fill="x")

        self.appt_canvas = tk.Canvas(p, bg=BG, highlightthickness=0)
        avsb = tk.Scrollbar(p, orient="vertical", command=self.appt_canvas.yview)
        self.appt_canvas.configure(yscrollcommand=avsb.set)
        avsb.pack(side="right", fill="y"); self.appt_canvas.pack(fill="both", expand=True)
        self.appt_body = tk.Frame(self.appt_canvas, bg=BG)
        self.appt_body_win = self.appt_canvas.create_window((0,0), window=self.appt_body, anchor="nw")
        self.appt_body.bind("<Configure>", lambda e: self.appt_canvas.configure(
            scrollregion=self.appt_canvas.bbox("all")))
        self.appt_canvas.bind("<Configure>", lambda e: self.appt_canvas.itemconfig(
            self.appt_body_win, width=e.width))
        self.appt_canvas.bind_all("<MouseWheel>", lambda e: self.appt_canvas.yview_scroll(
            int(-1*(e.delta/120)),"units") if self.active_view.get()=="appointments" else None)

    def _clear_appt_filters(self):
        self.appt_day_var.set("All"); self.appt_spec_var.set("All"); self._search_var.set("")
        self._refresh_appointments()

    def _refresh_appointments(self):
        for w in self.appt_body.winfo_children(): w.destroy()
        df = self.appt_day_var.get(); sf = self.appt_spec_var.get()
        q  = self._search_var.get().strip().lower()
        rows = [b for b in self.bookings
                if (df=="All" or b["day"]==df)
                and (sf=="All" or b["specialty"]==sf)
                and (not q or q in str(b.get("mrdno","")).lower()
                     or q in b.get("specialty","").lower()
                     or q in b.get("flow","").lower()
                     or q in b.get("doctor","").lower())]
        self.appt_count_lbl.config(text=f"{len(rows)} appointment{'s' if len(rows)!=1 else ''}")
        if not rows:
            wp = tk.Frame(self.appt_body, bg=BG); wp.pack(expand=True, fill="both", pady=80)
            tk.Label(wp, text="📭", font=("Segoe UI",28), bg=BG, fg=BORDER).pack()
            tk.Label(wp, text="No appointments found.", font=self.fnt_body, bg=BG, fg=TEXT_DIM).pack(pady=(8,0))
            return
        widths=[10,14,16,14,8,6,16,8,16,8]
        for i,b in enumerate(rows):
            rbg = CARD if i%2==0 else "#0f1419"
            row = tk.Frame(self.appt_body, bg=rbg); row.pack(fill="x")
            def _er(e,r=row,o=rbg):
                r.config(bg=CARD2)
                for ch in r.winfo_children():
                    try: ch.config(bg=CARD2)
                    except Exception: pass
            def _lr(e,r=row,o=rbg):
                r.config(bg=o)
                for ch in r.winfo_children():
                    try: ch.config(bg=o)
                    except Exception: pass
            row.bind("<Enter>",_er); row.bind("<Leave>",_lr)
            fc = {"Non-Dilated":TEAL,"Dilated":YELLOW,"Procedure":RED}.get(b.get("flow",""),TEXT)
            tw = clock_range(b.get("start_minute",0), b.get("duration_minutes",15))
            vals=[b.get("mrdno",""), b.get("gender","")+"  ·  "+b.get("age_cat",""),
                  b.get("specialty",""), b.get("flow",""), b.get("doctor",""),
                  b.get("day",""), tw, f"{b.get('duration_minutes',15)}m",
                  b.get("booked_at","")[:16]]
            fgs=[TEXT,TEXT_DIM,TEXT,fc,TEXT_DIM,ACCENT,TEXT,TEAL,TEXT_DIM]
            for v,w,fg in zip(vals,widths[:-1],fgs):
                lbl=tk.Label(row,text=str(v),font=self.fnt_small,bg=rbg,fg=fg,
                             width=w,anchor="w",pady=8,padx=8)
                lbl.pack(side="left"); lbl.bind("<Enter>",_er); lbl.bind("<Leave>",_lr)
            def _cancel(bid=b.get("id"), bdata=b):
                tw2=clock_range(bdata.get("start_minute",0),bdata.get("duration_minutes",15))
                pair_id = bdata.get("dilation_pair_id")
                if pair_id:
                    msg2 = f"Cancel BOTH dilation slots for {bdata['mrdno']} on {bdata['day']}?"
                else:
                    msg2 = f"Cancel {bdata['mrdno']} on {bdata['day']} {tw2}?"
                if messagebox.askyesno("Cancel", msg2):
                    if pair_id:
                        self.bookings=[x for x in self.bookings if x.get("dilation_pair_id")!=pair_id]
                        log.info("Cancelled dilation pair pair_id=%s",pair_id)
                    else:
                        self.bookings=[x for x in self.bookings if x.get("id")!=bid]
                        log.info("Cancelled booking id=%s",bid)
                    save_bookings(self.bookings); self._last_save=datetime.now()
                    self.toast.show(f"Cancelled: {bdata['mrdno']} on {bdata['day']}","warn")
                    self._refresh_appointments()
            cb=tk.Label(row,text=" Cancel ",font=self.fnt_small,bg="#3a1a1a",fg=RED,
                        cursor="hand2",padx=4,pady=3,highlightbackground="#6b2020",highlightthickness=1)
            cb.pack(side="left",padx=6)
            cb.bind("<Button-1>",lambda e,fn=_cancel:fn())
            cb.bind("<Enter>",lambda e,b=cb:b.config(bg="#5a1a1a",fg="#ff7070"))
            cb.bind("<Leave>",lambda e,b=cb:b.config(bg="#3a1a1a",fg=RED))
            tk.Frame(self.appt_body,bg=BORDER2,height=1).pack(fill="x")

    def _export_csv(self):
        if not self.bookings: self.toast.show("No bookings to export.","warn"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv"),("All","*.*")],
            initialfile=f"opd_bookings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        if not path: return
        keys=["id","mrdno","gender","age_cat","specialty","visit_cat","flow",
              "doctor","day","start_minute","duration_minutes","end_minute","pred_min","booked_at"]
        with open(path,"w",newline="",encoding="utf-8") as f:
            w=csv.DictWriter(f,fieldnames=keys,extrasaction="ignore")
            w.writeheader(); w.writerows(self.bookings)
        log.info("Exported %d bookings to %s",len(self.bookings),path)
        self.toast.show(f"Exported {len(self.bookings)} bookings.","success")

    def _focus_search(self):
        if self.active_view.get()!="appointments": self._show_view("appointments")
        self.after(100, self._appt_search.focus_set)

    # ── Data load ─────────────────────────────────────────────────────
    def _load_data_async(self):
        def _work():
            try:
                self.lookup = load_patient_lookup()
                self.pkg    = load_model()
                # Pre-compute global consult averages per specialty for new patients
                if "avg_consult" in self.lookup.columns:
                    self._global_consult_avg = (
                        self.lookup.groupby("Specialty")["avg_consult"]
                        .mean().dropna().round().astype(int).to_dict()
                    )
                else:
                    self._global_consult_avg = {}
                self._stop_loading_pulse(); self._stop_dot_animation()
                self.lbl_status_dot.config(fg=GREEN)
                self.lbl_status.config(text=" Ready", fg=GREEN)
                self._flash_ready()
                log.info("Data loaded successfully")
                self.toast.show("Data loaded. System ready.","success")
            except Exception as ex:
                self._stop_loading_pulse(); self._stop_dot_animation()
                self.lbl_status_dot.config(fg=YELLOW)
                self.lbl_status.config(text=" Demo mode", fg=YELLOW)
                log.warning("Data load failed: %s", ex)
                self.toast.show("Demo mode — no EHR data loaded.","warn",5000)
        threading.Thread(target=_work, daemon=True).start()

    # ── Look-up ───────────────────────────────────────────────────────
    def _lookup(self):
        mrdno = self.entry_mrdno.get().strip()
        if not mrdno: self.toast.show("Please enter a MRDNO.","warn"); return
        if self.lookup is None: self.toast.show("Data still loading…","warn"); return
        row = self.lookup.loc[self.lookup.index==mrdno]
        if row.empty: self.toast.show(f"MRDNO {mrdno} not found.","error"); return
        r = row.iloc[0]; g = str(r.get("Gender",""))
        self.patient_data = {
            "mrdno":       mrdno,
            "gender":      "Female" if g.startswith("Fe") or g.startswith("F") else "Male",
            "age_cat":     str(r.get("Age category","31-45")),
            "specialty":   str(r.get("Specialty","Retina")).strip(),
            "visit_cat":   str(r.get("Patient visit category","MRE")),
            "flow":        str(r.get("Patient Flow type","Non-Dilated")).strip(),
            "consultant":  str(r.get("Consultant Name","")),
            "designation": str(r.get("Consultant Designation","Specialist")),
            "session":     str(r.get("Session","Forenoon")),
            "arrival":     str(r.get("Arrival hour","09:00 - 10:00")),
            "past_visits":  int(r.get("past_visits",1)),
            "avg_workup":   float(r.get("avg_workup",60.0)),
            # Store all three flow-specific consult avgs so _render_patient can pick the right one
            "avg_consult_nondilated":  float(r["avg_consult_nondilated"])  if pd.notna(r.get("avg_consult_nondilated"))  else None,
            "avg_consult_dilated":     float(r["avg_consult_dilated"])     if pd.notna(r.get("avg_consult_dilated"))     else None,
            "avg_consult_procedure":   float(r["avg_consult_procedure"])   if pd.notna(r.get("avg_consult_procedure"))   else None,
        }
        log.info("Looked up MRDNO=%s", mrdno)
        self._render_patient()

    # ═══════════════════════════════════════════════════════════════════
    # Render patient
    # ═══════════════════════════════════════════════════════════════════
    def _render_patient(self):
        for w in self.content.winfo_children(): w.destroy()
        p   = self.patient_data
        if not p: return
        day = self.selected_day.get()
        existing = is_already_booked(self.bookings, p["mrdno"], day)

        # Patient card
        pc_o = tk.Frame(self.content, bg=BORDER, pady=1)
        pc_o.pack(fill="x", padx=16, pady=(14,4))
        pc = tk.Frame(pc_o, bg=CARD, pady=14, padx=16); pc.pack(fill="x")
        av_col = ACCENT2 if p["gender"]=="Male" else PURPLE
        tk.Label(pc, text=p["gender"][0], font=self.fnt_head,
                 bg=av_col, fg=TEXT, width=3, pady=8).pack(side="left", padx=(0,14))
        inf = tk.Frame(pc, bg=CARD); inf.pack(side="left", fill="x", expand=True)
        tk.Label(inf, text=f"MRDNO: {p['mrdno']}   ·   {p['specialty']}",
                 font=self.fnt_head, bg=CARD, fg=TEXT).pack(anchor="w")
        tk.Label(inf, text=(f"{p['gender']}   ·   Age {p['age_cat']}   ·   "
                            f"{p['past_visits']} past visits   ·   "
                            f"Avg workup {int(p['avg_workup'])} min"),
                 font=self.fnt_small, bg=CARD, fg=TEXT_DIM).pack(anchor="w", pady=(3,0))
        bc = {"MRE":("#1f3a5c","#79c0ff"),"SRE":("#1f3a1f",TEAL),
              "REG":(YELLOW_BG,YELLOW)}.get(p["visit_cat"],(CARD,TEXT_DIM))
        tk.Label(pc, text=f"  {p['visit_cat']}  ", font=self.fnt_tag,
                 bg=bc[0], fg=bc[1], padx=2, pady=4).pack(side="right", padx=(8,0))
        tk.Label(pc, text=f"  {p['designation']}  ", font=self.fnt_tag,
                 bg=CARD2, fg=TEXT_DIM, padx=2, pady=4).pack(side="right", padx=(0,6))

        # Already-booked banner
        if existing:
            tw = clock_range(existing.get("start_minute",0), existing.get("duration_minutes",15))
            bo = tk.Frame(self.content, bg=YELLOW, pady=1)
            bo.pack(fill="x", padx=16, pady=(4,0))
            bi = tk.Frame(bo, bg="#3a2000", pady=10); bi.pack(fill="x")
            tk.Label(bi, text=f"⚠   Already booked on {day}  —  {existing['doctor']}  ·  {tw}  ({existing.get('duration_minutes',15)} min)",
                     font=self.fnt_body, bg="#3a2000", fg=YELLOW).pack(side="left", padx=14)
            lnk = tk.Label(bi, text="View →", font=self.fnt_small, bg="#3a2000", fg=ACCENT, cursor="hand2")
            lnk.pack(side="right", padx=14)
            lnk.bind("<Button-1>", lambda e: self._show_view("appointments"))
            lnk.bind("<Enter>", lambda e: lnk.config(fg=TEXT))
            lnk.bind("<Leave>", lambda e: lnk.config(fg=ACCENT))

        # Predicted total stay by flow type (informational only — not used for slot sizing)
        self._section(self.content, "PREDICTED HOSPITAL STAY BY FLOW TYPE  ·  (total time, not slot length)", ACCENT2)
        predictions = {}
        for ft in FLOW_TYPES:
            predictions[ft] = (predict_workup(self.pkg, p["specialty"], ft, p["consultant"],
                                              p["session"], p["arrival"], p["avg_workup"])
                               if self.pkg else p["avg_workup"])
        ff = tk.Frame(self.content, bg=BG); ff.pack(fill="x", padx=16, pady=(0,4))
        for ft in FLOW_TYPES:
            self._flow_card(ff, ft, predictions[ft], ft.strip()==p["flow"].strip())

        pred_min = predictions.get(p["flow"].strip(), list(predictions.values())[0])
        self.patient_data["pred_min"] = pred_min

        # Consultation slot = patient's personal avg for TODAY's specific flow type
        # Must go through get_consult_slot so CONSULT_SLOT_CAP is applied
        # (Dilated EHR avg is the old 38-44 min full visit; cap clamps it to 6 min initial check)
        consult_min = get_consult_slot(p, p["flow"])
        self.patient_data["consult_min"] = consult_min

        # Info row: stay vs consult
        info_row = tk.Frame(self.content, bg=CARD2, padx=14, pady=8,
                            highlightbackground=BORDER2, highlightthickness=1)
        info_row.pack(fill="x", padx=16, pady=(0,4))
        tk.Label(info_row, text="🏥  Expected hospital stay:", font=self.fnt_small,
                 bg=CARD2, fg=TEXT_DIM).pack(side="left")
        tk.Label(info_row, text=f"  {int(pred_min)} min  ({fmt_range(pred_min)})",
                 font=self.fnt_small, bg=CARD2, fg=TEXT).pack(side="left")
        tk.Label(info_row, text="       🩺  Consultation slot booked:",
                 font=self.fnt_small, bg=CARD2, fg=TEXT_DIM).pack(side="left")
        tk.Label(info_row, text=f"  {consult_min} min  (avg from EHR sign-in/out)",
                 font=self.fnt_small, bg=CARD2, fg=TEAL).pack(side="left")

        # Ranked slots use consultation duration, not total workup
        self._ranked = [] if existing else rank_slots_v2(
            p["specialty"], day, p["flow"], self.bookings,
            consult_min, top_n=3, visit_cat=p["visit_cat"])

        # Recommendation banner
        sec_txt = f"  TIMELINE  —  {day.upper()}   ·   {p['specialty'].upper()}"
        if datetime.now().strftime("%a")==day and DAY_START_MIN <= datetime.now().hour*60+datetime.now().minute < DAY_END_MIN:
            sec_txt += f"   ·   ● LIVE {datetime.now().strftime('%H:%M')}"
        self._section(self.content, sec_txt, GREEN2)

        if self._ranked:
            best = self._ranked[0]; bd = best["breakdown"]
            zone = bd.get("zone",""); zc = ZONE_FG.get(zone, YELLOW)
            is_pref = bd.get("preferred", False)
            plan_b_txt = f"Plan B zone: {zone} ✓" if is_pref else f"⚠ Outside preferred window"
            rec_o = tk.Frame(self.content, bg=zc if is_pref else RED, pady=1)
            rec_o.pack(fill="x", padx=16, pady=(8,0))
            rec_bg = ZONE_COLOR.get(zone,"#2a2300") if is_pref else "#3a1a1a"
            rec = tk.Frame(rec_o, bg=rec_bg, pady=9); rec.pack(fill="x")
            li = tk.Frame(rec, bg=rec_bg); li.pack(side="left", padx=14)
            if best.get("is_dilation_pair"):
                slot_txt = (f"⚡  Recommended:  {best['doc']}   ·   "
                            f"Check {best['clock_range']} ({best['duration_minutes']} min)  "
                            f"→ wait {DILATION_WAIT_MIN} min →  "
                            f"Exam {best['exam_clock_range']} ({best['exam_duration_minutes']} min)")
            else:
                slot_txt = f"⚡  Recommended:  {best['doc']}   ·   {best['clock_range']}  ({best['duration_minutes']} min)"
            tk.Label(li, text=slot_txt, font=self.fnt_body, bg=rec_bg, fg=zc).pack(anchor="w")
            tk.Label(li, text=(f"  {plan_b_txt}   ·   "
                               f"Doc load: {bd['doc_booked_min']}/{bd['doc_capacity_min']} min   ·   "
                               f"Score: {best['score']:.3f}"),
                     font=self.fnt_small, bg=rec_bg, fg=TEXT_DIM).pack(anchor="w", pady=(2,0))
            bk = tk.Label(rec, text="Book this →", font=self.fnt_small,
                          bg=rec_bg, fg=ACCENT, cursor="hand2")
            bk.pack(side="right", padx=14)
            bk.bind("<Button-1>", lambda e, c=best: self._book_slot(
                c["doc"], c["start_minute"], c["duration_minutes"],
                c.get("exam_start_minute"), c.get("exam_duration_minutes")))
            bk.bind("<Enter>", lambda e: bk.config(fg=TEXT))
            bk.bind("<Leave>", lambda e: bk.config(fg=ACCENT))

        # Timeline view for each doctor
        self._timeline_view(p["specialty"], day, bool(existing), pred_min)

        # Legend
        leg = tk.Frame(self.content, bg=BG); leg.pack(fill="x", padx=16, pady=(8,2))
        for text, col in [("■ OPD gap",TEAL),("■ Booked",GREEN2),("■ OT",OT_COL),
                          ("■ Lunch",YELLOW_BG),("■ NA",BORDER2),
                          ("★#1 ★★#2 ★★★#3",YELLOW)]:
            tk.Label(leg, text=text, font=self.fnt_small, bg=BG, fg=col).pack(side="left", padx=(0,12))
        zl = tk.Frame(self.content, bg=BG); zl.pack(fill="x", padx=16, pady=(0,16))
        for zn,zc in [("Early",GREEN),("Mid",ACCENT),("Lunch",YELLOW),("Afternoon",PURPLE)]:
            tk.Label(zl, text=f"■ {zn}", font=self.fnt_small, bg=BG, fg=zc).pack(side="left", padx=(0,10))

    def _flow_card(self, parent, flow_type, pred_min, selected):
        bc = ACCENT if selected else BORDER2
        bg = "#1c2333" if selected else CARD
        o  = tk.Frame(parent, bg=bc, padx=1, pady=1)
        o.pack(side="left", fill="both", expand=True, padx=5)
        c  = tk.Frame(o, bg=bg, padx=12, pady=12); c.pack(fill="both", expand=True)
        il = tk.Label(c, text=FLOW_ICONS.get(flow_type,"👁"), font=("Segoe UI",16), bg=bg, fg=TEXT)
        il.pack(anchor="w")
        tl = tk.Label(c, text=flow_type, font=self.fnt_small, bg=bg, fg=ACCENT if selected else TEXT_DIM)
        tl.pack(anchor="w", pady=(2,0))
        vl = tk.Label(c, text=mins_to_hm(pred_min), font=self.fnt_big, bg=bg, fg=TEXT)
        vl.pack(anchor="w")
        rl = tk.Label(c, text=fmt_range(pred_min), font=self.fnt_small, bg=bg, fg=TEXT_DIM)
        rl.pack(anchor="w", pady=(0,4))
        if selected:
            tk.Label(c, text="▲ Selected", font=self.fnt_tag, bg=bg, fg=ACCENT).pack(anchor="w")
        all_ch=[c,il,tl,vl,rl]
        def _sel(): self.patient_data["flow"]=flow_type; self._render_patient()
        def _ent(e):
            c.config(bg=CARD2)
            for ch in all_ch:
                try: ch.config(bg=CARD2)
                except Exception: pass
            o.config(bg=ACCENT)
        def _lv(e):
            c.config(bg=bg)
            for ch in all_ch:
                try: ch.config(bg=bg)
                except Exception: pass
            o.config(bg=bc)
        for w in [o,c]+all_ch:
            w.bind("<Button-1>", lambda e: _sel())
            w.bind("<Enter>",    _ent)
            w.bind("<Leave>",    _lv)
            w.config(cursor="hand2")

    # ── Timeline view ─────────────────────────────────────────────────
    def _timeline_view(self, spec, day, already_booked, pred_min):
        """
        Draw a horizontal timeline bar per doctor.
        Width represents 08:00–17:00 (540 min total).
        Each segment is coloured by period status (OPD/OT/LUNCH/NA).
        Existing bookings are overlaid as darker filled blocks.
        Recommended slots are highlighted with rank stars.
        """
        sched = SCHEDULE.get(spec,{}).get(day,[])
        if not sched:
            tk.Label(self.content, text="No schedule for this specialty/day.",
                     font=self.fnt_body, bg=BG, fg=TEXT_DIM).pack(padx=16); return

        spec_prefix_tl = spec.strip().split()[0][:3].upper()
        rank_map = {(r["doc"],r["start_minute"]): i for i,r in enumerate(self._ranked)}

        outer = tk.Frame(self.content, bg=BORDER2, padx=1, pady=1)
        outer.pack(fill="x", padx=16, pady=(4,4))
        wrap = tk.Frame(outer, bg=CARD); wrap.pack(fill="x")

        # Time ruler
        ruler = tk.Frame(wrap, bg=CARD2, height=22)
        ruler.pack(fill="x")
        ruler_canvas = tk.Canvas(ruler, bg=CARD2, highlightthickness=0, height=22)
        ruler_canvas.pack(fill="x")
        ruler_canvas.bind("<Configure>", lambda e, rc=ruler_canvas: self._draw_ruler(rc))

        now_min = datetime.now().hour*60 + datetime.now().minute

        for di, doc_slots in enumerate(sched):
            doc_name  = f"{spec_prefix_tl}-Doc{di+1}"
            doc_bk    = [b for b in self.bookings if b["doctor"]==doc_name and b["day"]==day]
            total_opd = total_opd_minutes(doc_slots)
            booked_m  = sum(b["duration_minutes"] for b in doc_bk)
            pct       = round(booked_m/total_opd*100) if total_opd else 0
            bar_col   = GREEN if pct<60 else YELLOW if pct<85 else RED

            row = tk.Frame(wrap, bg=CARD); row.pack(fill="x", pady=2)

            # Doctor label column
            lbl_col = tk.Frame(row, bg=CARD, width=86)
            lbl_col.pack(side="left"); lbl_col.pack_propagate(False)
            tk.Label(lbl_col, text=doc_name, font=self.fnt_tag,
                     bg=CARD, fg=TEXT, anchor="w").pack(fill="x", padx=6, pady=(4,0))
            tk.Label(lbl_col, text=f"{booked_m}/{total_opd}m  {pct}%",
                     font=self.fnt_small, bg=CARD, fg=bar_col, anchor="w").pack(fill="x", padx=6)

            # Timeline canvas
            tl = tk.Canvas(row, bg=CARD2, height=48, highlightthickness=0)
            tl.pack(side="left", fill="x", expand=True, padx=(0,6), pady=4)

            # Draw after widget is sized
            tl.bind("<Configure>", lambda e, tl=tl, ds=doc_slots, bk=doc_bk,
                                          dn=doc_name, ab=already_booked, rm=rank_map:
                    self._draw_timeline(e.widget, e.width, ds, bk, dn, ab, rm, now_min, pred_min))

            # Hour separator
            tk.Frame(wrap, bg=BORDER2, height=1).pack(fill="x")

    def _draw_ruler(self, canvas):
        canvas.delete("all")
        w = canvas.winfo_width()
        if w < 10: return
        for h in range(8, 18):
            x = int((h*60 - DAY_START_MIN) / DAY_SPAN_MIN * w)
            canvas.create_line(x, 0, x, 22, fill=BORDER, width=1)
            canvas.create_text(x+2, 11, text=f"{h:02d}:00", fill=TEXT_DIM,
                               font=("Consolas",8), anchor="w")

    def _draw_timeline(self, canvas, w, doc_slots, doc_bk, doc_name,
                       already_booked, rank_map, now_min, pred_min):
        canvas.delete("all")
        if w < 10: return
        H = 48
        def xpos(m): return int((m - DAY_START_MIN) / DAY_SPAN_MIN * w)

        # Period backgrounds
        period_colors = {"OPD":"#1a2a1a","OT":"#1a1a2a","LUNCH":"#2a2000","NA":CARD2}
        for i, (ps, pe) in enumerate(zip(PERIOD_START_MIN, PERIOD_END_MIN)):
            status = doc_slots[i] if i<len(doc_slots) else "NA"
            col = period_colors.get(status, CARD2)
            canvas.create_rectangle(xpos(ps), 0, xpos(pe), H, fill=col, outline="")

        # Zone tint on OPD periods
        zone_tints = {"Early":"#1a3a1f","Mid":"#1c2d3a","Lunch":"#3a2a00","Afternoon":"#1f1f3a"}
        for i, (ps, pe) in enumerate(zip(PERIOD_START_MIN, PERIOD_END_MIN)):
            if i<len(doc_slots) and doc_slots[i]=="OPD":
                z = ZONE_LABEL.get(i,"")
                col = zone_tints.get(z, "#1a2a1a")
                canvas.create_rectangle(xpos(ps), 0, xpos(pe), H, fill=col, outline="")

        # Period dividers
        for ps in PERIOD_START_MIN:
            x = xpos(ps)
            canvas.create_line(x, 0, x, H, fill=BORDER2, width=1)

        # Recommended slots (highlight gaps before bookings overlay)
        if not already_booked:
            rank_labels = {0:"★ #1",1:"★★ #2",2:"★★★ #3"}
            rank_fgs    = {0:"#ffd700",1:"#c0c0c0",2:"#cd7f32"}
            for c in self._ranked:
                if c["doc"] != doc_name: continue
                rank = rank_map.get((doc_name, c["start_minute"]))
                if rank is None: continue
                x1 = xpos(c["start_minute"]); x2 = xpos(c["start_minute"]+c["duration_minutes"])
                col = {0:ACCENT2,1:"#1a3a2a",2:"#2a2a1a"}.get(rank,"#2a2a2a")
                canvas.create_rectangle(x1, 4, x2, H-4, fill=col, outline=rank_fgs.get(rank,ACCENT))
                mid = (x1+x2)//2
                canvas.create_text(mid, H//2, text=rank_labels.get(rank,""), fill=rank_fgs.get(rank,ACCENT),
                                   font=("Segoe UI",8,"bold"), anchor="center")

        # Existing bookings
        for b in doc_bk:
            x1 = xpos(b["start_minute"]); x2 = xpos(b["end_minute"])
            canvas.create_rectangle(x1, 8, x2, H-8, fill=GREEN2, outline=GREEN)
            mid = (x1+x2)//2
            dur_lbl = f"{b.get('duration_minutes',15)}m"
            canvas.create_text(mid, H//2, text=dur_lbl, fill=TEXT,
                               font=("Consolas",8), anchor="center")

        # "Now" indicator
        if DAY_START_MIN <= now_min < DAY_END_MIN:
            xn = xpos(now_min)
            canvas.create_line(xn, 0, xn, H, fill=RED, width=2, dash=(4,2))
            canvas.create_text(xn+2, 4, text=min_to_clock(now_min), fill=RED,
                               font=("Consolas",7), anchor="nw")

        # Click-to-book on free OPD areas
        if not already_booked:
            for c in self._ranked:
                if c["doc"] != doc_name: continue
                x1 = xpos(c["start_minute"]); x2 = xpos(c["start_minute"]+c["duration_minutes"])
                canvas.tag_bind(canvas.create_rectangle(x1, 0, x2, H, fill="", outline=""),
                                "<Button-1>",
                                lambda e, cc=c: self._book_slot(
                                    cc["doc"], cc["start_minute"], cc["duration_minutes"],
                                    cc.get("exam_start_minute"), cc.get("exam_duration_minutes")))
                canvas.tag_bind(canvas.create_rectangle(x1, 0, x2, H, fill="", outline=""),
                                "<Enter>",
                                lambda e, tl=canvas, a=x1, b_=x2:
                                    tl.create_rectangle(a,0,b_,H,fill=ACCENT2,outline="",tags="hover"))
                canvas.tag_bind(canvas.create_rectangle(x1, 0, x2, H, fill="", outline=""),
                                "<Leave>", lambda e, tl=canvas: tl.delete("hover"))

    # ── Booking actions ───────────────────────────────────────────────
    def _recommend_slot(self):
        if not self.patient_data: self.toast.show("Look up a patient first.","warn"); return
        p   = self.patient_data; day = self.selected_day.get()
        if is_already_booked(self.bookings, p["mrdno"], day):
            self.toast.show(f"Already booked on {day}.","warn"); return
        pred = p.get("pred_min",60) if self.pkg else p["avg_workup"]
        ranked = rank_slots_v2(p["specialty"],day,p["flow"],self.bookings,pred,top_n=3,visit_cat=p["visit_cat"])
        if not ranked: self.toast.show(f"No free slots for {p['specialty']} on {day}.","warn"); return
        best=ranked[0]; bd=best["breakdown"]; zone=bd.get("zone","")
        pref=bd.get("preferred",False)
        pref_wins=_get_preferred_periods(p["visit_cat"],p["flow"])
        wd={0:"8–9am",1:"9–10am",2:"10–11am",3:"11am–12pm",4:"12–1pm",5:"1–2:30pm",6:"2:30–3:30pm",7:"3:30–5pm"}
        pref_str=", ".join(wd.get(w,"") for w in pref_wins[:3])
        contrib={"Plan B window":f"{'✓ Preferred' if pref else '⚠ Outside'} ({zone})",
                 "Doctor load":  f"{bd['doc_booked_min']}/{bd['doc_capacity_min']} min booked",
                 "Congestion":   f"{bd['congestion']*100:.0f}% period filled",
                 "Dil penalty":  "Applied" if bd.get("dil_pen",0)>0 else "None",
                 "Lunch penalty":"Applied" if bd.get("lunch_pen",0)>0 else "None"}
        why="\n".join(f"   {k:<18}: {v}" for k,v in contrib.items())
        msg=(f"Plan B Recommended Slot\n{'─'*38}\n"
             f"Patient  : {p['visit_cat']}  ·  {p['flow']}\n"
             f"Preferred: {pref_str}\n\n"
             f"Doctor   : {best['doc']}\n"
             f"Time     : {best['clock_range']}  ({best['duration_minutes']} min)\n"
             f"Zone     : {zone}\nScore    : {best['score']:.3f}\n\n"
             f"Why:\n{why}\n\nBook it now?")
        if messagebox.askyesno("⚡ Recommended Slot", msg):
            self._book_slot(best["doc"], best["start_minute"], best["duration_minutes"],
                            best.get("exam_start_minute"), best.get("exam_duration_minutes"))
        else:
            self._render_patient()

    def _book_slot(self, doc, start_min, duration_min, exam_start=None, exam_dur=None):
        p   = self.patient_data; day = self.selected_day.get()
        existing = is_already_booked(self.bookings, p["mrdno"], day)
        if existing:
            tw=clock_range(existing.get("start_minute",0),existing.get("duration_minutes",15))
            self.toast.show(f"Already booked on {day}: {tw}","warn"); return
        conflict = has_interval_overlap(self.bookings, doc, day, start_min, duration_min)
        if conflict:
            tw=clock_range(conflict.get("start_minute",0),conflict.get("duration_minutes",15))
            self.toast.show(f"Slot conflicts with existing booking at {tw}","error"); return
        tw = clock_range(start_min, duration_min)
        is_pair = (exam_start is not None and exam_dur is not None)
        if is_pair:
            exam_tw = clock_range(exam_start, exam_dur)
            msg=(f"Confirm Dilated Booking (2 slots)\n\n"
                 f"MRDNO     : {p['mrdno']}\n"
                 f"Patient   : {p['gender']}  ·  {p['age_cat']}\n"
                 f"Specialty : {p['specialty']}\nFlow      : Dilated\n"
                 f"Doctor    : {doc}\nDay       : {day}\n\n"
                 f"Slot 1 (check)  : {tw}  ({duration_min} min)\n"
                 f"  ↓ wait {DILATION_WAIT_MIN} min for drops\n"
                 f"Slot 2 (exam)   : {exam_tw}  ({exam_dur} min)\n\nConfirm both?")
        else:
            msg=(f"Confirm Booking\n\n"
                 f"MRDNO     : {p['mrdno']}\n"
                 f"Patient   : {p['gender']}  ·  {p['age_cat']}\n"
                 f"Specialty : {p['specialty']}\nFlow      : {p['flow']}\n"
                 f"Doctor    : {doc}\nDay       : {day}\n"
                 f"Time      : {tw}  ({duration_min} min)\n\nConfirm?")
        if messagebox.askyesno("Confirm Booking", msg):
            now_iso = datetime.now().isoformat()
            base_id = f"{p['mrdno']}_{day}_{start_min}_{doc}".replace(" ","_")
            if is_pair:
                pair_id = f"DIL_{base_id}"
                b_check = {"id": base_id,
                           "mrdno":p["mrdno"],"gender":p["gender"],"age_cat":p["age_cat"],
                           "specialty":p["specialty"],"visit_cat":p["visit_cat"],"flow":p["flow"],
                           "doctor":doc,"day":day,"start_minute":start_min,
                           "duration_minutes":duration_min,"end_minute":start_min+duration_min,
                           "pred_min":round(duration_min,1),"booked_at":now_iso,
                           "dilation_pair_id":pair_id,"dilation_phase":"check"}
                b_exam  = {"id": f"{base_id}_exam",
                           "mrdno":p["mrdno"],"gender":p["gender"],"age_cat":p["age_cat"],
                           "specialty":p["specialty"],"visit_cat":p["visit_cat"],"flow":p["flow"],
                           "doctor":doc,"day":day,"start_minute":exam_start,
                           "duration_minutes":exam_dur,"end_minute":exam_start+exam_dur,
                           "pred_min":round(exam_dur,1),"booked_at":now_iso,
                           "dilation_pair_id":pair_id,"dilation_phase":"exam"}
                self.bookings.extend([b_check, b_exam])
                log.info("Booked dilated pair mrdno=%s doc=%s day=%s check=%s exam=%s",
                         p["mrdno"],doc,day,tw,exam_tw)
                self.toast.show(f"Booked (dilated): {p['mrdno']}  →  {doc}  |  {tw} + {exam_tw}","success")
            else:
                booking={"id":base_id,
                         "mrdno":p["mrdno"],"gender":p["gender"],"age_cat":p["age_cat"],
                         "specialty":p["specialty"],"visit_cat":p["visit_cat"],"flow":p["flow"],
                         "doctor":doc,"day":day,"start_minute":start_min,
                         "duration_minutes":duration_min,"end_minute":start_min+duration_min,
                         "pred_min":round(duration_min,1),"booked_at":now_iso}
                self.bookings.append(booking)
                log.info("Booked mrdno=%s doc=%s day=%s time=%s",p["mrdno"],doc,day,tw)
                self.toast.show(f"Booked: {p['mrdno']}  →  {doc}  |  {day}  {tw}","success")
            save_bookings(self.bookings); self._last_save=datetime.now()
            self._render_patient()

    # ── Manual entry ──────────────────────────────────────────────────
    def _open_manual(self):
        dlg = tk.Toplevel(self); dlg.title("Manual Patient Entry")
        dlg.configure(bg=BG); dlg.geometry("440x580"); dlg.grab_set(); dlg.resizable(False,False)
        hf = tk.Frame(dlg, bg=CARD2, pady=12); hf.pack(fill="x")
        tk.Label(hf, text="✏   Manual Patient Entry", font=self.fnt_head, bg=CARD2, fg=TEXT).pack(side="left",padx=16)
        tk.Frame(dlg, bg=ACCENT2, height=2).pack(fill="x")
        body = tk.Frame(dlg, bg=BG); body.pack(fill="both", expand=True, padx=20, pady=8)

        def lbl(t): tk.Label(body,text=t,font=self.fnt_small,bg=BG,fg=TEXT_DIM).pack(anchor="w",pady=(8,2))
        def combo(var,choices):
            cb=ttk.Combobox(body,textvariable=var,values=choices,font=self.fnt_body,state="readonly")
            cb.pack(fill="x"); return cb
        def entry(var):
            wr=tk.Frame(body,bg=BORDER,padx=1,pady=1); wr.pack(fill="x")
            ent=tk.Entry(wr,textvariable=var,font=self.fnt_body,bg=CARD,fg=TEXT,
                         insertbackground=ACCENT,relief="flat",bd=4); ent.pack(fill="x")
            ent.bind("<FocusIn>",  lambda e: wr.config(bg=ACCENT))
            ent.bind("<FocusOut>", lambda e: wr.config(bg=BORDER))
            return ent

        v_mrdno=tk.StringVar(master=dlg,value="MANUAL"); v_gender=tk.StringVar(master=dlg,value="Male")
        v_age=tk.StringVar(master=dlg,value="31-45");    v_spec=tk.StringVar(master=dlg,value="Retina")
        v_flow=tk.StringVar(master=dlg,value="Non-Dilated"); v_vcat=tk.StringVar(master=dlg,value="MRE")
        v_avg=tk.StringVar(master=dlg,value="—")

        lbl("MRDNO / Label");  entry(v_mrdno)
        lbl("Gender");         combo(v_gender,["Male","Female"])
        lbl("Age Category");   combo(v_age,AGE_CATS)
        lbl("Specialty");      cb_spec=combo(v_spec,SPECIALTIES)
        lbl("Flow Type");      cb_flow=combo(v_flow,FLOW_TYPES)
        lbl("Visit Category"); combo(v_vcat,["MRE","SRE","REG"])

        # Predicted total stay (model output — informational)
        lbl("Predicted Hospital Stay (min)  — informational")
        pw=tk.Frame(body,bg=BORDER,padx=1,pady=1); pw.pack(fill="x")
        pe=tk.Entry(pw,textvariable=v_avg,font=self.fnt_body,bg=CARD2,fg=TEXT_DIM,
                    relief="flat",bd=4,state="readonly"); pe.pack(side="left",fill="x",expand=True)
        pl=tk.Label(pw,text="auto",font=self.fnt_small,bg=CARD2,fg=TEXT_DIM,padx=8); pl.pack(side="right")

        # Consultation duration — what actually gets booked
        v_consult = tk.StringVar(master=dlg, value="15")
        lbl("Consultation Slot (min)  — this is what gets booked on the doctor's calendar")
        cw=tk.Frame(body,bg=BORDER,padx=1,pady=1); cw.pack(fill="x")
        cent=tk.Entry(cw,textvariable=v_consult,font=self.fnt_body,bg=CARD,fg=TEAL,
                      insertbackground=ACCENT,relief="flat",bd=4); cent.pack(side="left",fill="x",expand=True)
        tk.Label(cw,text="min",font=self.fnt_small,bg=CARD,fg=TEXT_DIM,padx=8).pack(side="right")
        cent.bind("<FocusIn>",  lambda _: cw.config(bg=ACCENT))
        cent.bind("<FocusOut>", lambda _: cw.config(bg=BORDER))

        def _pred(*_):
            if not self.pkg: v_avg.set("—"); pl.config(text="no model",fg=TEXT_DIM); return
            try:
                pred=predict_workup(self.pkg,v_spec.get(),v_flow.get(),"","Forenoon","09:00 - 10:00",0.0)
                v_avg.set(str(int(pred))); pl.config(text="✓",fg=TEXT_DIM)
            except Exception: v_avg.set("—"); pl.config(text="err",fg=RED)
        cb_spec.bind("<<ComboboxSelected>>",_pred); cb_flow.bind("<<ComboboxSelected>>",_pred)
        dlg.after(100,_pred)

        def apply():
            try: avg=float(v_avg.get())
            except ValueError: avg=60.0
            try: consult=max(5, int(v_consult.get()))
            except ValueError: consult=15
            self.patient_data={"mrdno":v_mrdno.get(),"gender":v_gender.get(),"age_cat":v_age.get(),
                               "specialty":v_spec.get(),"visit_cat":v_vcat.get(),"flow":v_flow.get(),
                               "consultant":"","designation":"Specialist","session":"Forenoon",
                               "arrival":"09:00 - 10:00","past_visits":1,
                               "avg_workup":avg,"avg_consult":float(consult)}
            dlg.destroy(); self._render_patient()

        bf=tk.Frame(dlg,bg=BG); bf.pack(fill="x",padx=20,pady=12)
        ab=tk.Label(bf,text="  Show Schedule →  ",font=self.fnt_body,bg=ACCENT2,fg=TEXT,
                    cursor="hand2",padx=6,pady=7,highlightbackground=ACCENT,highlightthickness=1)
        ab.pack(side="right")
        ab.bind("<Button-1>",lambda e:apply()); ab.bind("<Enter>",lambda e:ab.config(bg=ACCENT,fg=BG))
        ab.bind("<Leave>",lambda e:ab.config(bg=ACCENT2,fg=TEXT))
        cl=tk.Label(bf,text="  Cancel  ",font=self.fnt_small,bg=CARD,fg=TEXT_DIM,
                    cursor="hand2",padx=4,pady=7,highlightbackground=BORDER,highlightthickness=1)
        cl.pack(side="right",padx=(0,8))
        cl.bind("<Button-1>",lambda e:dlg.destroy()); cl.bind("<Enter>",lambda e:cl.config(bg=CARD2,fg=TEXT))
        cl.bind("<Leave>",lambda e:cl.config(bg=CARD,fg=TEXT_DIM))


# ── Entry point ────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = OPDSchedulerApp()
    style = ttk.Style(app)
    style.theme_use("clam")
    style.configure("TCombobox",
        fieldbackground=CARD, background=CARD, foreground=TEXT,
        selectbackground=ACCENT2, selectforeground=TEXT,
        arrowcolor=TEXT_DIM, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
    style.map("TCombobox",
        fieldbackground=[("readonly",CARD)],
        foreground=[("readonly",TEXT)],
        background=[("readonly",CARD)])
    app.mainloop()
