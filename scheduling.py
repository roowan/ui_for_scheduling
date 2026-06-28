"""
OPD Workup Scheduler — Tkinter GUI
Files needed (same folder):
  workup_model_age7_ar1.pkl, ehr_age7_ar1.xlsx,
  Consultant_schedule.docx, patient_lookup.csv (auto-built)
"""

import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont
import pandas as pd
import numpy as np
import pickle
import os
import json
import threading
from datetime import datetime

# ── Paths ─────────────────────────────────────────────────────────────
BASE         = os.path.dirname(os.path.abspath(__file__))
PKL_PATH     = os.path.join(BASE, "workup_model_age7_ar1.pkl")
EHR_PATH     = os.path.join(BASE, "ehr_age7_ar1.xlsx")
LOOKUP_CACHE = os.path.join(BASE, "patient_lookup.csv")
BOOKINGS_FILE= os.path.join(BASE, "bookings.json")

# ── Palette ───────────────────────────────────────────────────────────
BG       = "#0d1117"
CARD     = "#161b22"
CARD2    = "#1c2333"
BORDER   = "#30363d"
BORDER2  = "#21262d"
ACCENT   = "#58a6ff"
ACCENT2  = "#1f6feb"
GREEN    = "#3fb950"
GREEN_LT = "#1a4a27"
GREEN2   = "#238636"
OT_COL   = "#6e7681"
TEXT     = "#e6edf3"
TEXT_DIM = "#8b949e"
RED      = "#f85149"
TEAL     = "#39d353"
YELLOW   = "#d29922"
YELLOW_BG= "#3a2a00"
SIDEBAR  = "#010409"
SIDEBAR2 = "#0d1117"

# ── Broad period definitions (used to interpret SCHEDULE data) ─────────
SLOT_LABELS = ["08-09","09-10","10-11","11-12","12-13:30","13:30-14:30","14:30-15:30","15:30-17"]
SLOT_START_H= [8, 9, 10, 11, 12, 13.5, 14.5, 15.5]
SLOT_DUR_H  = [1, 1,  1,  1,  1.5,  1,    1,    1.5]

# ── 15-minute consultation slots ───────────────────────────────────────
def _gen_15min_labels():
    labels = []
    t = 8 * 60
    while t < 17 * 60:
        h, m = divmod(t, 60)
        h2, m2 = divmod(t + 15, 60)
        labels.append(f"{h:02d}:{m:02d}–{h2:02d}:{m2:02d}")
        t += 15
    return labels

SLOT_15MIN = _gen_15min_labels()

def get_broad_idx(slot_15_label):
    """Maps a '08:15–08:30' label to its broad period index (0-7)."""
    start_str = slot_15_label.split("–")[0]
    h, m = map(int, start_str.split(":"))
    start_min = h * 60 + m
    for i, (sh, dur) in enumerate(zip(SLOT_START_H, SLOT_DUR_H)):
        if int(sh * 60) <= start_min < int((sh + dur) * 60):
            return i
    return -1

def get_slot_status_15min(slot_15_label, doc_slots):
    """Returns OPD/OT/LUNCH/NA for a 15-min slot label."""
    idx = get_broad_idx(slot_15_label)
    if 0 <= idx < len(doc_slots):
        return doc_slots[idx]
    return "NA"

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

def _lerp_color(c1, c2, t):
    """Linearly interpolate between two hex colors, t in [0,1]."""
    r1,g1,b1 = int(c1[1:3],16), int(c1[3:5],16), int(c1[5:7],16)
    r2,g2,b2 = int(c2[1:3],16), int(c2[3:5],16), int(c2[5:7],16)
    r = int(r1 + (r2-r1)*t)
    g = int(g1 + (g2-g1)*t)
    b = int(b1 + (b2-b1)*t)
    return f"#{r:02x}{g:02x}{b:02x}"

# ── Booking store ─────────────────────────────────────────────────────
def load_bookings():
    if os.path.exists(BOOKINGS_FILE):
        with open(BOOKINGS_FILE) as f:
            return json.load(f)
    return []

def save_bookings(bookings):
    with open(BOOKINGS_FILE, "w") as f:
        json.dump(bookings, f, indent=2)

def is_already_booked(bookings, mrdno, day):
    """Returns the existing booking if patient already booked on that day."""
    for b in bookings:
        if str(b["mrdno"]) == str(mrdno) and b["day"] == day:
            return b
    return None

def is_slot_booked(bookings, doctor, day, slot_15_label):
    """Returns the booking if this 15-min slot is already taken, else None."""
    for b in bookings:
        if b["doctor"] == doctor and b["day"] == day and b["slot"] == slot_15_label:
            return b
    return None

# ── Data loading ──────────────────────────────────────────────────────
def load_patient_lookup():
    if os.path.exists(LOOKUP_CACHE):
        return pd.read_csv(LOOKUP_CACHE, dtype={"MRDNO": str}).set_index("MRDNO")
    df = pd.read_excel(EHR_PATH)
    df["MRDNO"] = df["MRDNO"].astype(str)
    df["workup_min"] = df["TOTAL_WORKUP_TIME"].apply(
        lambda x: x.hour*60+x.minute if hasattr(x,"hour") else None)
    df_sorted = df.sort_values("VISITDATE", ascending=False)
    info = df_sorted.groupby("MRDNO").first()[[
        "Gender","Age category","Specialty","Patient visit category",
        "Patient Flow type","Consultant Name","Consultant Designation",
        "Day","Session","Arrival hour"]]
    vc  = df.groupby("MRDNO").size().rename("past_visits")
    avg = df.groupby("MRDNO")["workup_min"].mean().rename("avg_workup")
    lookup = info.join(vc).join(avg)
    lookup.reset_index().to_csv(LOOKUP_CACHE, index=False)
    return lookup

def load_model():
    class SafeUnpickler(pickle.Unpickler):
        def find_class(self, module, name):
            if "xgboost" in module:
                class Dummy: pass
                return Dummy
            return super().find_class(module, name)
    with open(PKL_PATH, "rb") as f:
        return SafeUnpickler(f).load()

def predict_workup(pkg, spec, flow, consultant, session, arrival, avg_workup):
    em  = pkg["encoding_maps"]
    s   = em["spec_mean"].get(spec.strip(),  em["global_mean"])
    fl  = em["flow_mean"].get(flow.strip(),  em["global_mean"])
    c   = em["consultant_mean"].get(consultant, em["global_mean"])
    bl  = s*0.3 + fl*0.5 + c*0.2
    if session == "Afternoon": bl *= 0.95
    ah  = em["arrival_map"].get(arrival, 10)
    if ah >= 14: bl *= 1.05
    elif ah <= 9: bl *= 0.95
    return round(bl*0.6 + avg_workup*0.4, 1)

# ── Research-backed slot recommender — Plan B (Lin, Jin & Chia 2014 adapted) ──
#
# From our EHR analysis (n=225,961):
#   Procedure follow-up  : 120 min avg workup  → must start early (8–10am)
#   Dilated follow-up    :  72 min avg workup  → early-mid (8–12pm)
#   Non-Dilated follow-up:  40 min avg workup  → flexible, mid-session (10am–3pm)
#   Walk-in REG          :  79 min avg workup  → two dedicated windows: 9–11am + 2–4pm
#
# Waiting time by hour from EHR:
#   8–9am  : 16.0 min  ← lowest, ideal for heavy cases
#   10–11am: 19.9 min  ← moderate
#   1pm    : 27.0 min  ← lunch surge peak (avoid for long cases)
#   3–4pm  : 16.5 min  ← drops back, good for walk-ins

def _slot_start_min(slot_15_label):
    """Convert a slot label like '09:15–09:30' to minutes from midnight."""
    h, m = map(int, slot_15_label.split("\u2013")[0].split(":"))
    return h * 60 + m
#
# Plan B scheduling windows (matching paper's Table IV structure):
PLAN_B_WINDOWS = {
    # (visit_cat, flow_type) → list of preferred broad period indices [0-7]
    # Broad periods: 0=8-9, 1=9-10, 2=10-11, 3=11-12, 4=12-13:30, 5=13:30-14:30, 6=14:30-15:30, 7=15:30-17
    ("SRE", "Procedure"):     [0, 1],            # 8–10am — longest, must start early
    ("MRE", "Procedure"):     [0, 1],
    ("REG", "Procedure"):     [0, 1],
    ("SRE", "Dilated"):       [0, 1, 2, 3],      # 8–12pm — needs dilation wait time
    ("MRE", "Dilated"):       [0, 1, 2, 3],
    ("REG", "Dilated"):       [0, 1, 2, 3],
    ("SRE", "Non-Dilated"):   [2, 3, 4, 6, 7],   # 10am–3pm — shortest, most flexible
    ("MRE", "Non-Dilated"):   [2, 3, 4, 6, 7],
    ("REG", "Non-Dilated"):   [1, 2, 6, 7],      # Walk-in Non-Dilated: 9–11am + 2:30–5pm
}
# Walk-in REG fallback (any flow): two windows per paper Plan B
WALKIN_WINDOWS = [1, 2, 6, 7]   # 9–11am + 2:30–5pm

# Zone labels for UI display
ZONE_LABEL = {
    0: "Early",    1: "Early",
    2: "Mid",      3: "Mid",      4: "Mid",
    5: "Lunch",
    6: "Afternoon",7: "Afternoon",
}
ZONE_COLOR = {
    "Early":     "#1a3a1f",   # dark green
    "Mid":       "#1c2d3a",   # dark blue
    "Lunch":     "#3a2a00",   # dark amber
    "Afternoon": "#1f1f3a",   # dark purple
}
ZONE_FG = {
    "Early": "#3fb950", "Mid": "#58a6ff",
    "Lunch": "#d29922", "Afternoon": "#8957e5",
}

def _get_preferred_periods(visit_cat, flow):
    """Return list of preferred broad period indices for this patient type."""
    key = (visit_cat, flow.strip())
    if key in PLAN_B_WINDOWS:
        return PLAN_B_WINDOWS[key]
    if visit_cat == "REG":
        return WALKIN_WINDOWS
    return list(range(8))   # no restriction — any OPD slot

def score_slot(slot_15, doc_name, day, flow, bookings, visit_cat="MRE"):
    """
    Return (score, breakdown_dict). Lower score = better recommendation.
    Scoring is based on:
      1. Plan B window match (0 = preferred period, 1 = non-preferred)
      2. Doctor load balance
      3. Period congestion (avoid overfull periods)
      4. Dilation penalty (Dilated must not land in late afternoon)
    """
    broad = get_broad_idx(slot_15)

    # 1. Plan B window match — hard preference
    preferred = _get_preferred_periods(visit_cat, flow)
    plan_b_score = 0.0 if broad in preferred else 0.6

    # 2. Doctor load
    doc_count = sum(1 for b in bookings
                    if b["day"] == day and b["doctor"] == doc_name)
    doc_load = min(1.0, doc_count / 36.0)

    # 3. Period congestion (all doctors in this period)
    period_count = sum(1 for b in bookings
                       if b["day"] == day
                       and get_broad_idx(b.get("slot","")) == broad)
    congestion = min(1.0, period_count / 8.0)

    # 4. Dilation penalty — dilation drop takes 20-30 min so avoid tail of day
    t_norm = (_slot_start_min(slot_15) - 8*60) / (9*60)
    dil_pen = 0.4 if (flow.strip() == "Dilated" and t_norm > 0.70) else 0.0

    # 5. Lunch hour penalty (1pm spike = 27 min wait from EHR analysis)
    lunch_pen = 0.3 if broad == 5 else 0.0

    # 6. Walk-in buffer: reserve at least 2 slots per period for walk-ins
    # If this is a follow-up patient, penalise periods already saturated by follow-ups
    fu_in_period = sum(1 for b in bookings
                       if b["day"] == day
                       and get_broad_idx(b.get("slot","")) == broad
                       and b.get("visit_cat","") in ("SRE","MRE"))
    walkin_pen = 0.2 if fu_in_period >= 6 else 0.0

    score = (0.40 * plan_b_score
             + 0.25 * doc_load
             + 0.15 * congestion
             + 0.10 * dil_pen
             + 0.05 * lunch_pen
             + 0.05 * walkin_pen)

    return score, {
        "plan_b":    plan_b_score,
        "doc_load":  doc_load,
        "congestion":congestion,
        "dil_pen":   dil_pen,
        "lunch_pen": lunch_pen,
        "preferred": broad in preferred,
        "zone":      ZONE_LABEL.get(broad, ""),
    }

def rank_slots(spec, day, flow, bookings, top_n=3, visit_cat="MRE"):
    """Score every free OPD slot for a specialty/day/visit_cat.
    Returns sorted list of {doc, slot, score, breakdown}, best first."""
    sched = SCHEDULE.get(spec, {}).get(day, [])
    candidates = []
    for di, doc_slots in enumerate(sched):
        doc_name = f"Doc {di+1}"
        for slot_15 in SLOT_15MIN:
            if get_slot_status_15min(slot_15, doc_slots) != "OPD":
                continue
            if is_slot_booked(bookings, doc_name, day, slot_15):
                continue
            score, breakdown = score_slot(
                slot_15, doc_name, day, flow, bookings, visit_cat)
            candidates.append({"doc": doc_name, "slot": slot_15,
                               "score": score, "breakdown": breakdown})
    candidates.sort(key=lambda c: (c["score"], _slot_start_min(c["slot"]), c["doc"]))
    return candidates[:top_n] if top_n else candidates

# ═════════════════════════════════════════════════════════════════════
# Main App
# ═════════════════════════════════════════════════════════════════════
class OPDSchedulerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OPD Workup Scheduler")
        self.geometry("1280x880")
        self.configure(bg=BG)
        self.resizable(True, True)

        self.patient_data  = None
        self.pkg           = None
        self.lookup        = None
        self.bookings      = load_bookings()
        self.selected_day  = tk.StringVar(master=self, value=self._today_day())
        self.active_view   = tk.StringVar(master=self, value="scheduler")

        # Animation state
        self._loading_dots  = 0
        self._pulse_dir     = 1
        self._pulse_val     = 0
        self._loading_job   = None
        self._dot_job       = None

        self._build_fonts()
        self._build_ui()
        self._load_data_async()

    def _today_day(self):
        d = datetime.now().strftime("%a")
        return d if d in DAYS_ORDER else "Mon"

    def _build_fonts(self):
        self.fnt_title  = tkfont.Font(family="Segoe UI", size=18, weight="bold")
        self.fnt_head   = tkfont.Font(family="Segoe UI", size=12, weight="bold")
        self.fnt_body   = tkfont.Font(family="Segoe UI", size=11)
        self.fnt_small  = tkfont.Font(family="Segoe UI", size=9)
        self.fnt_big    = tkfont.Font(family="Segoe UI", size=26, weight="bold")
        self.fnt_label  = tkfont.Font(family="Segoe UI", size=10)
        self.fnt_nav    = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        self.fnt_mono   = tkfont.Font(family="Consolas",  size=10)
        self.fnt_tag    = tkfont.Font(family="Segoe UI", size=8,  weight="bold")

    # ── Animation helpers ─────────────────────────────────────────────
    def _animate_hover_in(self, widget, bg_from, bg_to, fg_from, fg_to, step=0, steps=8):
        t = step / steps
        widget.config(bg=_lerp_color(bg_from, bg_to, t),
                      fg=_lerp_color(fg_from, fg_to, t))
        if step < steps:
            widget._hover_job = self.after(
                16, lambda: self._animate_hover_in(
                    widget, bg_from, bg_to, fg_from, fg_to, step+1, steps))

    def _animate_hover_out(self, widget, bg_from, bg_to, fg_from, fg_to, step=0, steps=8):
        t = step / steps
        widget.config(bg=_lerp_color(bg_from, bg_to, t),
                      fg=_lerp_color(fg_from, fg_to, t))
        if step < steps:
            widget._hover_job = self.after(
                16, lambda: self._animate_hover_out(
                    widget, bg_from, bg_to, fg_from, fg_to, step+1, steps))

    def _cancel_hover(self, widget):
        if hasattr(widget, "_hover_job") and widget._hover_job:
            self.after_cancel(widget._hover_job)
            widget._hover_job = None

    def _start_loading_pulse(self):
        """Pulses the status label colour while loading."""
        colors = [TEXT_DIM, YELLOW, "#e6a817", YELLOW, TEXT_DIM]
        self._pulse_val = (self._pulse_val + 1) % len(colors)
        try:
            self.lbl_status.config(fg=colors[self._pulse_val])
        except Exception:
            return
        self._loading_job = self.after(400, self._start_loading_pulse)

    def _stop_loading_pulse(self):
        if self._loading_job:
            self.after_cancel(self._loading_job)
            self._loading_job = None

    def _start_dot_animation(self, label, base="Loading"):
        """Animates '…' dots on a label."""
        self._loading_dots = (self._loading_dots + 1) % 4
        dots = "." * self._loading_dots
        try:
            label.config(text=base + dots)
        except Exception:
            return
        self._dot_job = self.after(500, lambda: self._start_dot_animation(label, base))

    def _stop_dot_animation(self):
        if self._dot_job:
            self.after_cancel(self._dot_job)
            self._dot_job = None

    def _flash_ready(self):
        """Brief green flash on the status label when data loads."""
        colors = [GREEN, "#5fff78", GREEN, "#2ea043", GREEN]
        def step(i=0):
            if i < len(colors):
                try:
                    self.lbl_status.config(fg=colors[i])
                except Exception:
                    return
                self.after(120, lambda: step(i+1))
        step()

    # ── UI skeleton ───────────────────────────────────────────────────
    def _build_ui(self):
        # ── Sidebar ──────────────────────────────────────────────────
        self.sidebar = tk.Frame(self, bg=SIDEBAR, width=196)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Logo area with accent bar on the left edge
        logo_wrap = tk.Frame(self.sidebar, bg=SIDEBAR)
        logo_wrap.pack(fill="x", pady=(0,0))

        accent_bar = tk.Frame(logo_wrap, bg=ACCENT, width=3)
        accent_bar.pack(side="left", fill="y", pady=18)

        logo_inner = tk.Frame(logo_wrap, bg=SIDEBAR)
        logo_inner.pack(side="left", fill="both", expand=True, padx=(10,0))

        tk.Label(logo_inner, text="🏥  OPD",
                 font=self.fnt_title, bg=SIDEBAR, fg=TEXT).pack(pady=(18,2), anchor="w")
        tk.Label(logo_inner, text="Workup Scheduler",
                 font=self.fnt_small, bg=SIDEBAR, fg=TEXT_DIM).pack(anchor="w", pady=(0,14))

        # Divider
        tk.Frame(self.sidebar, bg=BORDER2, height=1).pack(fill="x", padx=12)
        tk.Label(self.sidebar, text="NAVIGATION",
                 font=self.fnt_tag, bg=SIDEBAR, fg="#444c56").pack(
                 anchor="w", padx=20, pady=(14,6))

        self._nav_btn("📋   Scheduler",    "scheduler")
        self._nav_btn("📅   Appointments", "appointments")

        # Bottom status area
        tk.Frame(self.sidebar, bg=BORDER2, height=1).pack(side="bottom", fill="x", padx=12, pady=(0,4))
        status_wrap = tk.Frame(self.sidebar, bg=SIDEBAR)
        status_wrap.pack(side="bottom", fill="x", padx=14, pady=(0,16))

        dot_frame = tk.Frame(status_wrap, bg=SIDEBAR)
        dot_frame.pack(anchor="w")
        self.lbl_status_dot = tk.Label(dot_frame, text="●", font=self.fnt_small,
                                        bg=SIDEBAR, fg=YELLOW)
        self.lbl_status_dot.pack(side="left")
        self.lbl_status = tk.Label(dot_frame, text="Loading",
                                    font=self.fnt_small, bg=SIDEBAR, fg=YELLOW)
        self.lbl_status.pack(side="left", padx=(4,0))

        self._start_dot_animation(self.lbl_status)
        self._start_loading_pulse()

        # ── Main area ─────────────────────────────────────────────────
        self.main = tk.Frame(self, bg=BG)
        self.main.pack(side="left", fill="both", expand=True)

        # Pages
        self.page_scheduler    = tk.Frame(self.main, bg=BG)
        self.page_appointments = tk.Frame(self.main, bg=BG)

        self._build_scheduler_page()
        self._build_appointments_page()
        self._show_view("scheduler")

    def _nav_btn(self, label, view):
        btn = tk.Label(self.sidebar, text=label, font=self.fnt_nav,
                       bg=SIDEBAR, fg=TEXT_DIM, anchor="w",
                       padx=20, pady=11, cursor="hand2")
        btn.pack(fill="x")
        btn._view = view
        btn._hover_job = None

        def on_enter(e, b=btn):
            self._cancel_hover(b)
            is_active = self.active_view.get() == b._view
            if not is_active:
                self._animate_hover_in(b, SIDEBAR, CARD2, TEXT_DIM, TEXT)

        def on_leave(e, b=btn):
            self._cancel_hover(b)
            is_active = self.active_view.get() == b._view
            if not is_active:
                self._animate_hover_out(b, CARD2, SIDEBAR, TEXT, TEXT_DIM)

        btn.bind("<Enter>",    on_enter)
        btn.bind("<Leave>",    on_leave)
        btn.bind("<Button-1>", lambda e, v=view: self._show_view(v))

        if not hasattr(self, "_nav_btns"): self._nav_btns = []
        self._nav_btns.append(btn)

    def _show_view(self, view):
        self.active_view.set(view)
        self.page_scheduler.pack_forget()
        self.page_appointments.pack_forget()
        if view == "scheduler":
            self.page_scheduler.pack(fill="both", expand=True)
        else:
            self.page_appointments.pack(fill="both", expand=True)
            self._refresh_appointments()
        for b in getattr(self, "_nav_btns", []):
            if b._view == view:
                b.config(bg=CARD2, fg=ACCENT)
            else:
                b.config(bg=SIDEBAR, fg=TEXT_DIM)

    # ── Scheduler page ────────────────────────────────────────────────
    def _build_scheduler_page(self):
        p = self.page_scheduler

        # ── Top bar ───────────────────────────────────────────────────
        top = tk.Frame(p, bg=CARD, pady=0)
        top.pack(fill="x")

        inner = tk.Frame(top, bg=CARD, pady=10)
        inner.pack(fill="x", padx=16)

        # Label + entry group
        mrd_group = tk.Frame(inner, bg=CARD2, padx=2, pady=2,
                             highlightbackground=BORDER, highlightthickness=1)
        mrd_group.pack(side="left")
        tk.Label(mrd_group, text=" MRDNO ", font=self.fnt_small,
                 bg="#1f6feb", fg=TEXT, padx=6).pack(side="left")
        self.entry_mrdno = tk.Entry(mrd_group, font=self.fnt_head, width=14,
                                     bg=CARD2, fg=TEXT, insertbackground=ACCENT,
                                     relief="flat", bd=4)
        self.entry_mrdno.pack(side="left", padx=(4,6))
        self.entry_mrdno.bind("<Return>", lambda e: self._lookup())
        self.entry_mrdno.bind("<FocusIn>",
            lambda e: mrd_group.config(highlightbackground=ACCENT))
        self.entry_mrdno.bind("<FocusOut>",
            lambda e: mrd_group.config(highlightbackground=BORDER))

        # Day selector
        day_group = tk.Frame(inner, bg=CARD2, padx=2, pady=2,
                             highlightbackground=BORDER, highlightthickness=1)
        day_group.pack(side="left", padx=(10,0))
        tk.Label(day_group, text=" Day ", font=self.fnt_small,
                 bg="#2ea043", fg=TEXT, padx=6).pack(side="left")
        self.day_cb = ttk.Combobox(day_group, textvariable=self.selected_day,
                                    values=DAYS_ORDER, width=6,
                                    font=self.fnt_body, state="readonly")
        self.day_cb.pack(side="left", padx=(4,6))

        # Buttons
        self.btn_lookup = tk.Label(inner, text="  Look up  ", font=self.fnt_body,
                                    bg=ACCENT2, fg=TEXT, cursor="hand2",
                                    padx=4, pady=5,
                                    highlightbackground=ACCENT, highlightthickness=1)
        self.btn_lookup.pack(side="left", padx=(12,0))
        self.btn_lookup.bind("<Button-1>", lambda e: self._lookup())
        self.btn_lookup.bind("<Enter>",
            lambda e: self.btn_lookup.config(bg=ACCENT, fg=BG))
        self.btn_lookup.bind("<Leave>",
            lambda e: self.btn_lookup.config(bg=ACCENT2, fg=TEXT))

        self.btn_manual = tk.Label(inner, text="  Manual Entry  ", font=self.fnt_small,
                                    bg=CARD, fg=TEXT_DIM, cursor="hand2",
                                    padx=4, pady=5,
                                    highlightbackground=BORDER, highlightthickness=1)
        self.btn_manual.pack(side="left", padx=(8,0))
        self.btn_manual.bind("<Button-1>", lambda e: self._open_manual())
        self.btn_manual.bind("<Enter>",
            lambda e: self.btn_manual.config(bg=CARD2, fg=TEXT))
        self.btn_manual.bind("<Leave>",
            lambda e: self.btn_manual.config(bg=CARD,  fg=TEXT_DIM))

        self.btn_recommend = tk.Label(inner, text="  \u26a1 Recommend Slot  ",
                                      font=self.fnt_small, bg="#3a2f00", fg=YELLOW,
                                      cursor="hand2", padx=4, pady=5,
                                      highlightbackground=YELLOW, highlightthickness=1)
        self.btn_recommend.pack(side="left", padx=(8,0))
        self.btn_recommend.bind("<Button-1>", lambda e: self._recommend_slot())
        self.btn_recommend.bind("<Enter>",
            lambda e: self.btn_recommend.config(bg=YELLOW, fg=BG))
        self.btn_recommend.bind("<Leave>",
            lambda e: self.btn_recommend.config(bg="#3a2f00", fg=YELLOW))

        # Thin accent line under top bar
        tk.Frame(p, bg=ACCENT2, height=2).pack(fill="x")

        # ── Scrollable content ─────────────────────────────────────────
        self.canvas = tk.Canvas(p, bg=BG, highlightthickness=0)
        vsb = tk.Scrollbar(p, orient="vertical", command=self.canvas.yview,
                           troughcolor=CARD, bg=BORDER)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(fill="both", expand=True)

        self.content = tk.Frame(self.canvas, bg=BG)
        self.content_win = self.canvas.create_window(
            (0,0), window=self.content, anchor="nw")
        self.content.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            self.content_win, width=e.width))
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(
            int(-1*(e.delta/120)), "units"))

        # Placeholder with animated dots
        self._ph_wrap = tk.Frame(self.content, bg=BG)
        self._ph_wrap.pack(expand=True, fill="both", pady=100)
        tk.Label(self._ph_wrap, text="🔍", font=("Segoe UI", 32),
                 bg=BG, fg=BORDER).pack()
        self._ph_lbl = tk.Label(self._ph_wrap,
            text="Enter a MRDNO and click  Look up",
            font=self.fnt_body, bg=BG, fg=TEXT_DIM)
        self._ph_lbl.pack(pady=(8,0))
        tk.Label(self._ph_wrap,
            text="or use Manual Entry to enter patient details directly",
            font=self.fnt_small, bg=BG, fg="#444c56").pack(pady=(4,0))

    # ── Appointments page ─────────────────────────────────────────────
    def _build_appointments_page(self):
        p = self.page_appointments

        hdr = tk.Frame(p, bg=CARD, pady=0)
        hdr.pack(fill="x")
        hdr_inner = tk.Frame(hdr, bg=CARD, pady=12)
        hdr_inner.pack(fill="x", padx=16)
        tk.Label(hdr_inner, text="📅  All Appointments",
                 font=self.fnt_title, bg=CARD, fg=TEXT).pack(side="left")
        tk.Frame(p, bg=ACCENT2, height=2).pack(fill="x")

        # Filter bar
        fbar = tk.Frame(p, bg=BG, pady=10)
        fbar.pack(fill="x", padx=16)

        tk.Label(fbar, text="Day:", font=self.fnt_small,
                 bg=BG, fg=TEXT_DIM).pack(side="left")
        self.appt_day_var = tk.StringVar(master=self, value="All")
        day_filter = ttk.Combobox(fbar, textvariable=self.appt_day_var,
                                   values=["All"]+DAYS_ORDER, width=8,
                                   font=self.fnt_label, state="readonly")
        day_filter.pack(side="left", padx=(4,16))
        day_filter.bind("<<ComboboxSelected>>", lambda e: self._refresh_appointments())

        tk.Label(fbar, text="Specialty:", font=self.fnt_small,
                 bg=BG, fg=TEXT_DIM).pack(side="left")
        self.appt_spec_var = tk.StringVar(master=self, value="All")
        spec_filter = ttk.Combobox(fbar, textvariable=self.appt_spec_var,
                                    values=["All"]+SPECIALTIES, width=18,
                                    font=self.fnt_label, state="readonly")
        spec_filter.pack(side="left", padx=(4,16))
        spec_filter.bind("<<ComboboxSelected>>", lambda e: self._refresh_appointments())

        btn_clr = tk.Label(fbar, text="  Clear Filters  ", font=self.fnt_small,
                           bg=CARD, fg=TEXT_DIM, cursor="hand2",
                           padx=4, pady=3,
                           highlightbackground=BORDER, highlightthickness=1)
        btn_clr.pack(side="left")
        btn_clr.bind("<Button-1>", lambda e: self._clear_appt_filters())
        btn_clr.bind("<Enter>", lambda e: btn_clr.config(bg=CARD2, fg=TEXT))
        btn_clr.bind("<Leave>", lambda e: btn_clr.config(bg=CARD,  fg=TEXT_DIM))

        self.appt_count_lbl = tk.Label(fbar, text="", font=self.fnt_small,
                                        bg=BG, fg=TEXT_DIM)
        self.appt_count_lbl.pack(side="right")

        tk.Frame(p, bg=BORDER2, height=1).pack(fill="x")

        # Table header
        cols   = ["MRDNO","Patient","Specialty","Flow","Doctor","Day","Slot","Booked At","Action"]
        widths = [10, 14, 16, 14, 8, 6, 14, 18, 8]

        hrow = tk.Frame(p, bg=CARD2)
        hrow.pack(fill="x")
        for c, w in zip(cols, widths):
            tk.Label(hrow, text=c.upper(), font=self.fnt_tag,
                     bg=CARD2, fg="#444c56",
                     width=w, anchor="w", pady=10, padx=8).pack(side="left")

        tk.Frame(p, bg=ACCENT2, height=1).pack(fill="x")

        # Scrollable table body
        self.appt_canvas = tk.Canvas(p, bg=BG, highlightthickness=0)
        avsb = tk.Scrollbar(p, orient="vertical", command=self.appt_canvas.yview)
        self.appt_canvas.configure(yscrollcommand=avsb.set)
        avsb.pack(side="right", fill="y")
        self.appt_canvas.pack(fill="both", expand=True)

        self.appt_body = tk.Frame(self.appt_canvas, bg=BG)
        self.appt_body_win = self.appt_canvas.create_window(
            (0,0), window=self.appt_body, anchor="nw")
        self.appt_body.bind("<Configure>", lambda e: self.appt_canvas.configure(
            scrollregion=self.appt_canvas.bbox("all")))
        self.appt_canvas.bind("<Configure>", lambda e: self.appt_canvas.itemconfig(
            self.appt_body_win, width=e.width))

    def _clear_appt_filters(self):
        self.appt_day_var.set("All")
        self.appt_spec_var.set("All")
        self._refresh_appointments()

    def _refresh_appointments(self):
        for w in self.appt_body.winfo_children():
            w.destroy()

        day_f  = self.appt_day_var.get()
        spec_f = self.appt_spec_var.get()

        rows = [b for b in self.bookings
                if (day_f  == "All" or b["day"]       == day_f)
                and (spec_f == "All" or b["specialty"] == spec_f)]

        count = len(rows)
        self.appt_count_lbl.config(
            text=f"{count} appointment{'s' if count != 1 else ''}")

        if not rows:
            wrap = tk.Frame(self.appt_body, bg=BG)
            wrap.pack(expand=True, fill="both", pady=80)
            tk.Label(wrap, text="📭", font=("Segoe UI", 28),
                     bg=BG, fg=BORDER).pack()
            tk.Label(wrap, text="No appointments found.",
                     font=self.fnt_body, bg=BG, fg=TEXT_DIM).pack(pady=(8,0))
            return

        widths = [10, 14, 16, 14, 8, 6, 14, 18, 8]
        for i, b in enumerate(rows):
            row_bg = CARD if i % 2 == 0 else "#0f1419"
            row = tk.Frame(self.appt_body, bg=row_bg,
                           highlightbackground=row_bg, highlightthickness=0)
            row.pack(fill="x")

            # Row hover highlight
            def _enter_row(e, r=row, orig=row_bg):
                r.config(bg=CARD2)
                for ch in r.winfo_children():
                    try: ch.config(bg=CARD2)
                    except Exception: pass

            def _leave_row(e, r=row, orig=row_bg):
                r.config(bg=orig)
                for ch in r.winfo_children():
                    try: ch.config(bg=orig)
                    except Exception: pass

            row.bind("<Enter>", _enter_row)
            row.bind("<Leave>", _leave_row)

            flow_col = {
                "Non-Dilated": TEAL,
                "Dilated":     YELLOW,
                "Procedure":   "#f78166"
            }.get(b.get("flow",""), TEXT)

            vals = [
                b.get("mrdno",""),
                b.get("gender","") + "  ·  " + b.get("age_cat",""),
                b.get("specialty",""),
                b.get("flow",""),
                b.get("doctor",""),
                b.get("day",""),
                b.get("slot",""),
                b.get("booked_at","")[:16],
            ]
            fgs = [TEXT, TEXT_DIM, TEXT, flow_col, TEXT_DIM, ACCENT, TEXT, TEXT_DIM]

            for v, w, fg in zip(vals, widths[:-1], fgs):
                lbl = tk.Label(row, text=str(v), font=self.fnt_small,
                               bg=row_bg, fg=fg,
                               width=w, anchor="w", pady=9, padx=8)
                lbl.pack(side="left")
                lbl.bind("<Enter>", _enter_row)
                lbl.bind("<Leave>", _leave_row)

            # Cancel button
            def _cancel(bid=b.get("id"), rf=self._refresh_appointments):
                if messagebox.askyesno("Cancel Appointment",
                    f"Cancel appointment for MRDNO {b['mrdno']} on {b['day']} {b['slot']}?"):
                    self.bookings = [x for x in self.bookings if x.get("id") != bid]
                    save_bookings(self.bookings)
                    rf()

            cancel_btn = tk.Label(row, text=" Cancel ", font=self.fnt_small,
                                   bg="#3a1a1a", fg=RED, cursor="hand2",
                                   padx=4, pady=3,
                                   highlightbackground="#6b2020", highlightthickness=1)
            cancel_btn.pack(side="left", padx=6)
            cancel_btn.bind("<Button-1>", lambda e, fn=_cancel: fn())
            cancel_btn.bind("<Enter>",
                lambda e, b=cancel_btn: b.config(bg="#5a1a1a", fg="#ff7070"))
            cancel_btn.bind("<Leave>",
                lambda e, b=cancel_btn: b.config(bg="#3a1a1a", fg=RED))

            tk.Frame(self.appt_body, bg=BORDER2, height=1).pack(fill="x")

    # ── Data load ─────────────────────────────────────────────────────
    def _load_data_async(self):
        def _work():
            try:
                self.lookup = load_patient_lookup()
                self.pkg    = load_model()
                self._stop_loading_pulse()
                self._stop_dot_animation()
                self.lbl_status_dot.config(fg=GREEN)
                self.lbl_status.config(text=" Ready", fg=GREEN)
                self._flash_ready()
            except Exception as ex:
                self._stop_loading_pulse()
                self._stop_dot_animation()
                self.lbl_status_dot.config(fg=RED)
                self.lbl_status.config(text=f" Error", fg=RED)
        threading.Thread(target=_work, daemon=True).start()

    # ── Look-up ───────────────────────────────────────────────────────
    def _lookup(self):
        mrdno = self.entry_mrdno.get().strip()
        if not mrdno:
            messagebox.showwarning("Input", "Please enter a MRDNO.")
            return
        if self.lookup is None:
            messagebox.showinfo("Loading", "Data still loading, please wait.")
            return
        row = self.lookup.loc[self.lookup.index == mrdno]
        if row.empty:
            messagebox.showerror("Not Found", f"MRDNO {mrdno} not found.")
            return
        r = row.iloc[0]
        g = str(r.get("Gender",""))
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
            "past_visits": int(r.get("past_visits",1)),
            "avg_workup":  float(r.get("avg_workup",60.0)),
        }
        self._render_patient()

    # ── Render ────────────────────────────────────────────────────────
    def _render_patient(self):
        for w in self.content.winfo_children(): w.destroy()
        p   = self.patient_data
        day = self.selected_day.get()

        existing = is_already_booked(self.bookings, p["mrdno"], day)

        # ── Patient card ──────────────────────────────────────────────
        pc_outer = tk.Frame(self.content, bg=BORDER, pady=1)
        pc_outer.pack(fill="x", padx=16, pady=(14,4))
        pc = tk.Frame(pc_outer, bg=CARD, pady=14, padx=16)
        pc.pack(fill="x")

        av_col = "#1f6feb" if p["gender"] == "Male" else "#8957e5"
        av = tk.Label(pc, text=p["gender"][0], font=self.fnt_head,
                      bg=av_col, fg=TEXT, width=3, pady=8)
        av.pack(side="left", padx=(0,14))

        inf = tk.Frame(pc, bg=CARD)
        inf.pack(side="left", fill="x", expand=True)

        tk.Label(inf, text=f"MRDNO: {p['mrdno']}   ·   {p['specialty']}",
                 font=self.fnt_head, bg=CARD, fg=TEXT).pack(anchor="w")
        tk.Label(inf,
                 text=(f"{p['gender']}   ·   Age {p['age_cat']}   ·   "
                       f"{p['past_visits']} past visits   ·   "
                       f"Avg workup {int(p['avg_workup'])} min"),
                 font=self.fnt_small, bg=CARD, fg=TEXT_DIM).pack(anchor="w", pady=(3,0))

        badge_cfg = {
            "MRE": ("#1f3a5c", "#79c0ff"),
            "SRE": ("#1f3a1f", TEAL),
            "REG": ("#3a2a1f", YELLOW),
        }.get(p["visit_cat"], (CARD, TEXT_DIM))
        tk.Label(pc, text=f"  {p['visit_cat']}  ", font=self.fnt_tag,
                 bg=badge_cfg[0], fg=badge_cfg[1],
                 padx=2, pady=4).pack(side="right", padx=(8,0))

        # ── Already-booked banner ────────────────────────────────────
        if existing:
            banner_outer = tk.Frame(self.content, bg=YELLOW, pady=1)
            banner_outer.pack(fill="x", padx=16, pady=(4,0))
            banner = tk.Frame(banner_outer, bg="#3a2000", pady=10)
            banner.pack(fill="x")
            tk.Label(banner,
                     text=f"⚠   Already booked on {day}  —  {existing['doctor']}   ·   {existing['slot']}",
                     font=self.fnt_body, bg="#3a2000", fg=YELLOW).pack(side="left", padx=14)
            lnk = tk.Label(banner, text="View in Appointments →",
                           font=self.fnt_small, bg="#3a2000", fg=ACCENT, cursor="hand2")
            lnk.pack(side="right", padx=14)
            lnk.bind("<Button-1>", lambda e: self._show_view("appointments"))
            lnk.bind("<Enter>", lambda e: lnk.config(fg=TEXT))
            lnk.bind("<Leave>", lambda e: lnk.config(fg=ACCENT))

        # ── Flow cards ───────────────────────────────────────────────
        sec_lbl = tk.Frame(self.content, bg=BG)
        sec_lbl.pack(fill="x", padx=16, pady=(16,4))
        tk.Frame(sec_lbl, bg=ACCENT2, width=3).pack(side="left", fill="y")
        tk.Label(sec_lbl, text="  PREDICTED WORKUP TIME BY FLOW TYPE",
                 font=self.fnt_tag, bg=BG, fg=TEXT_DIM).pack(side="left")

        predictions = {}
        for ft in FLOW_TYPES:
            predictions[ft] = predict_workup(
                self.pkg, p["specialty"], ft, p["consultant"],
                p["session"], p["arrival"], p["avg_workup"])

        ff = tk.Frame(self.content, bg=BG)
        ff.pack(fill="x", padx=16, pady=(0,4))
        for ft in FLOW_TYPES:
            self._flow_card(ff, ft, predictions[ft], ft.strip()==p["flow"].strip())

        # ── Slot grid ─────────────────────────────────────────────────
        spec = p["specialty"]
        sec_lbl2 = tk.Frame(self.content, bg=BG)
        sec_lbl2.pack(fill="x", padx=16, pady=(16,4))
        tk.Frame(sec_lbl2, bg=GREEN2, width=3).pack(side="left", fill="y")
        tk.Label(sec_lbl2,
                 text=f"  AVAILABLE OPD SLOTS  —  {day.upper()}   ·   {spec.upper()}",
                 font=self.fnt_tag, bg=BG, fg=TEXT_DIM).pack(side="left")

        pred_min = predictions.get(p["flow"].strip(), list(predictions.values())[0])
        self.patient_data["pred_min"] = pred_min

        # Plan B ranking — pass visit_cat so scoring uses correct preferred windows
        self._ranked = [] if existing else rank_slots(
            spec, day, p["flow"], self.bookings, top_n=3, visit_cat=p["visit_cat"])
        self._rank_map = {(r["doc"], r["slot"]): i
                          for i, r in enumerate(self._ranked)}

        # Recommendation banner with Plan B zone explanation
        if self._ranked:
            best = self._ranked[0]
            bd   = best["breakdown"]
            zone = bd.get("zone", "")
            zone_col = ZONE_FG.get(zone, YELLOW)
            is_preferred = bd.get("preferred", False)
            plan_b_txt = (f"Plan B zone: {zone} ✓" if is_preferred
                          else f"⚠ Outside preferred {zone} window")

            rec_outer = tk.Frame(self.content, bg=zone_col if is_preferred else RED,
                                  pady=1)
            rec_outer.pack(fill="x", padx=16, pady=(8,0))
            rec_bg = ZONE_COLOR.get(zone, "#2a2300") if is_preferred else "#3a1a1a"
            rec = tk.Frame(rec_outer, bg=rec_bg, pady=9)
            rec.pack(fill="x")

            left_info = tk.Frame(rec, bg=rec_bg)
            left_info.pack(side="left", padx=14)
            tk.Label(left_info,
                text=(f"⚡  Recommended:  {best['doc']}   ·   {best['slot']}"),
                font=self.fnt_body, bg=rec_bg, fg=zone_col).pack(anchor="w")
            tk.Label(left_info,
                text=(f"  {plan_b_txt}   ·   "
                      f"Doctor load: {bd['doc_load']*36:.0f} booked   ·   "
                      f"Period: {bd['congestion']*8:.0f} busy   ·   "
                      f"Score: {best['score']:.3f}"),
                font=self.fnt_small, bg=rec_bg, fg=TEXT_DIM).pack(anchor="w", pady=(2,0))

            bk = tk.Label(rec, text="Book this →", font=self.fnt_small,
                          bg=rec_bg, fg=ACCENT, cursor="hand2")
            bk.pack(side="right", padx=14)
            bk.bind("<Button-1>",
                    lambda e, b=best: self._book_slot(b["doc"], b["slot"]))
            bk.bind("<Enter>", lambda e: bk.config(fg=TEXT))
            bk.bind("<Leave>", lambda e: bk.config(fg=ACCENT))

        self._slot_grid(spec, day, bool(existing))

        # Legend — zone colours + slot types
        legend = tk.Frame(self.content, bg=BG)
        legend.pack(fill="x", padx=16, pady=(8,4))
        for text, col in [("● Available", TEAL), ("● Booked", "#39d353"),
                           ("● OT", OT_COL), ("● Lunch/NA", BORDER)]:
            tk.Label(legend, text=text, font=self.fnt_small,
                     bg=BG, fg=col).pack(side="left", padx=(0,14))

        zone_legend = tk.Frame(self.content, bg=BG)
        zone_legend.pack(fill="x", padx=16, pady=(0,4))
        tk.Label(zone_legend, text="Plan B zones:", font=self.fnt_small,
                 bg=BG, fg=TEXT_DIM).pack(side="left", padx=(0,8))
        for zone_name, zone_col in [("Early 8–10am", "#3fb950"),
                                     ("Mid 10am–1pm", "#58a6ff"),
                                     ("Lunch 1pm", "#d29922"),
                                     ("Afternoon 2–5pm", "#8957e5")]:
            tk.Label(zone_legend, text=f"■ {zone_name}", font=self.fnt_small,
                     bg=BG, fg=zone_col).pack(side="left", padx=(0,12))

        tk.Label(self.content,
                 text="  ★ #1  ★★ #2  ★★★ #3 = recommended slots by Plan B score",
                 font=self.fnt_small, bg=BG, fg=TEXT_DIM).pack(
                 anchor="w", padx=16, pady=(0,16))

    def _flow_card(self, parent, flow_type, pred_min, selected):
        border_col = ACCENT if selected else BORDER2
        bg_col     = "#1c2333" if selected else CARD

        outer = tk.Frame(parent, bg=border_col, padx=1, pady=1)
        outer.pack(side="left", fill="both", expand=True, padx=5)

        card = tk.Frame(outer, bg=bg_col, padx=12, pady=12)
        card.pack(fill="both", expand=True)

        icon_lbl = tk.Label(card, text=FLOW_ICONS.get(flow_type,"👁"),
                             font=("Segoe UI", 16), bg=bg_col, fg=TEXT)
        icon_lbl.pack(anchor="w")

        type_lbl = tk.Label(card, text=flow_type, font=self.fnt_small,
                             bg=bg_col, fg=TEXT_DIM if not selected else ACCENT)
        type_lbl.pack(anchor="w", pady=(2,0))

        time_lbl = tk.Label(card, text=mins_to_hm(pred_min),
                             font=self.fnt_big, bg=bg_col, fg=TEXT)
        time_lbl.pack(anchor="w")

        range_lbl = tk.Label(card, text=fmt_range(pred_min),
                              font=self.fnt_small, bg=bg_col, fg=TEXT_DIM)
        range_lbl.pack(anchor="w", pady=(0,4))

        if selected:
            tk.Label(card, text="▲ Selected", font=self.fnt_tag,
                     bg=bg_col, fg=ACCENT).pack(anchor="w")

        all_children = [card, icon_lbl, type_lbl, time_lbl, range_lbl]

        def _select():
            self.patient_data["flow"] = flow_type
            self._render_patient()

        def _enter(e):
            card.config(bg=CARD2)
            for ch in all_children:
                try: ch.config(bg=CARD2)
                except Exception: pass
            outer.config(bg=ACCENT if not selected else ACCENT)

        def _leave(e):
            card.config(bg=bg_col)
            for ch in all_children:
                try: ch.config(bg=bg_col)
                except Exception: pass
            outer.config(bg=border_col)

        for w in [outer, card] + all_children:
            w.bind("<Button-1>", lambda e: _select())
            w.bind("<Enter>", _enter)
            w.bind("<Leave>", _leave)
            w.config(cursor="hand2")

    def _slot_grid(self, spec, day, already_booked):
        sched = SCHEDULE.get(spec, {}).get(day, [])
        if not sched:
            tk.Label(self.content, text="No schedule for this specialty/day.",
                     font=self.fnt_body, bg=BG, fg=TEXT_DIM).pack(padx=16)
            return

        num_docs = len(sched)
        outer = tk.Frame(self.content, bg=BORDER2, padx=1, pady=1)
        outer.pack(fill="x", padx=16)
        grid = tk.Frame(outer, bg=CARD)
        grid.pack(fill="x")

        # Header
        tk.Label(grid, text="TIME", font=self.fnt_tag, bg=CARD2, fg="#444c56",
                 width=14, anchor="w", pady=8, padx=8).grid(
            row=0, column=0, padx=1, pady=1, sticky="nsew")
        for di in range(num_docs):
            tk.Label(grid, text=f"DOCTOR {di+1}", font=self.fnt_tag,
                     bg=CARD2, fg="#444c56",
                     width=10, anchor="center", pady=8).grid(
                row=0, column=di+1, padx=1, pady=1, sticky="nsew")

        tk.Frame(grid, bg=ACCENT2, height=1).grid(
            row=1, column=0, columnspan=num_docs+1, sticky="ew")

        for ri, slot_15 in enumerate(SLOT_15MIN):
            row_bg = CARD if ri % 2 == 0 else "#0f1419"
            tk.Label(grid, text=slot_15, font=self.fnt_mono,
                     bg=row_bg, fg=TEXT_DIM,
                     width=14, anchor="w", pady=5, padx=8).grid(
                row=ri+2, column=0, padx=1, pady=1, sticky="nsew")
            for di, doc_slots in enumerate(sched):
                doc_name = f"Doc {di+1}"
                val  = get_slot_status_15min(slot_15, doc_slots)
                taken = is_slot_booked(self.bookings, doc_name, day, slot_15)
                rank = getattr(self, "_rank_map", {}).get((doc_name, slot_15))
                self._slot_cell(grid, ri+2, di+1, val, slot_15, doc_name,
                                already_booked, taken, row_bg, rank)

        grid.columnconfigure(0, weight=2)
        for di in range(num_docs):
            grid.columnconfigure(di+1, weight=1)

    def _slot_cell(self, parent, row, col, val, slot_15, doc_name,
                   already_booked, taken, row_bg, rank=None):
        if val in ("LUNCH", "NA"):
            text = "☕ LUNCH" if val == "LUNCH" else "—"
            fg   = "#7a6020" if val == "LUNCH" else "#333a44"
            tk.Label(parent, text=text, font=self.fnt_small, bg=row_bg,
                     fg=fg, width=10, anchor="center", pady=5).grid(
                row=row, column=col, padx=1, pady=1, sticky="nsew")
            return

        if val == "OT":
            tk.Label(parent, text="OT", font=self.fnt_small, bg=row_bg,
                     fg=OT_COL, width=10, anchor="center", pady=5).grid(
                row=row, column=col, padx=1, pady=1, sticky="nsew")
            return

        # Determine Plan B zone for this slot
        broad     = get_broad_idx(slot_15)
        zone      = ZONE_LABEL.get(broad, "")
        zone_fg   = ZONE_FG.get(zone, TEAL)
        zone_bg   = ZONE_COLOR.get(zone, "#1a2a1a")

        # val == "OPD"
        if taken:
            lbl = tk.Label(parent, text="✓ Booked", font=self.fnt_small,
                           bg="#1a3a1a", fg=TEAL, width=10, anchor="center",
                           pady=5, cursor="arrow")
            lbl.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")

        elif already_booked:
            lbl = tk.Label(parent, text="OPD", font=self.fnt_small,
                           bg="#1c2128", fg="#444c56", width=10, anchor="center",
                           pady=5, cursor="arrow")
            lbl.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")

        else:
            # Rank badges: ★ = best, ★★ = 2nd, ★★★ = 3rd
            rank_txt = {0: "★ #1", 1: "★★ #2", 2: "★★★ #3"}.get(rank, "OPD")
            rank_fg  = {0: "#ffd700", 1: "#c0c0c0", 2: "#cd7f32"}.get(rank, zone_fg)
            cell_bg  = zone_bg if rank is not None else row_bg
            cell_fg  = rank_fg if rank is not None else zone_fg

            lbl = tk.Label(parent, text=rank_txt, font=self.fnt_small,
                           bg=cell_bg, fg=cell_fg, width=10, anchor="center",
                           pady=5, cursor="hand2")
            lbl.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")
            lbl.bind("<Button-1>", lambda e, d=doc_name, s=slot_15: self._book_slot(d, s))
            hover_bg = ACCENT2 if rank is not None else zone_bg
            lbl.bind("<Enter>",    lambda e, w=lbl, hb=hover_bg: w.config(bg=hb, fg=TEXT))
            lbl.bind("<Leave>",    lambda e, w=lbl, ob=cell_bg, of=cell_fg: w.config(bg=ob, fg=of))

    def _recommend_slot(self):
        if self.patient_data is None:
            messagebox.showinfo("No Patient",
                "Look up a MRDNO or use Manual Entry first.")
            return
        p   = self.patient_data
        day = self.selected_day.get()
        if is_already_booked(self.bookings, p["mrdno"], day):
            messagebox.showinfo("Already Booked",
                f"MRDNO {p['mrdno']} already has a booking on {day}.")
            return
        ranked = rank_slots(p["specialty"], day, p["flow"], self.bookings,
                            top_n=3, visit_cat=p["visit_cat"])
        if not ranked:
            messagebox.showwarning("No Slots",
                f"No free OPD slots for {p['specialty']} on {day}.")
            return
        best = ranked[0]
        bd   = best["breakdown"]
        zone = bd.get("zone", "")
        preferred = bd.get("preferred", False)

        # Plan B window context per patient type
        preferred_windows = _get_preferred_periods(p["visit_cat"], p["flow"])
        window_desc = {
            0: "8–9am", 1: "9–10am", 2: "10–11am", 3: "11am–12pm",
            4: "12–1pm", 5: "1–2:30pm (Lunch)", 6: "2:30–3:30pm", 7: "3:30–5pm"
        }
        pref_str = ", ".join(window_desc.get(w,"") for w in preferred_windows[:3])

        contrib = {
            "Plan B window match": f"{'✓ In preferred zone' if preferred else '⚠ Outside preferred zone'} ({zone})",
            "Doctor load":         f"{bd['doc_load']*36:.0f} slots already booked",
            "Period congestion":   f"{bd['congestion']*8:.0f} patients in this period",
            "Dilation penalty":    f"{'Applied' if bd.get('dil_pen',0) > 0 else 'None'}",
            "Lunch surge penalty": f"{'Applied' if bd.get('lunch_pen',0) > 0 else 'None'}",
        }
        why = "\n".join(f"   {k:<22}: {v}" for k, v in contrib.items())
        msg = (f"Plan B Recommended Slot\n"
               f"{'─'*40}\n"
               f"Patient type : {p['visit_cat']}  ·  {p['flow']}\n"
               f"Preferred    : {pref_str}\n\n"
               f"Doctor    : {best['doc']}\n"
               f"Time Slot : {best['slot']}   [{zone}]\n"
               f"Score     : {best['score']:.3f}  (lower = better)\n\n"
               f"Why this slot:\n{why}\n\n"
               f"Book it now?")
        if messagebox.askyesno("⚡ Plan B Recommended Slot", msg):
            self._book_slot(best["doc"], best["slot"])
        else:
            self._render_patient()

    def _book_slot(self, doc, slot):
        p   = self.patient_data
        day = self.selected_day.get()

        existing = is_already_booked(self.bookings, p["mrdno"], day)
        if existing:
            messagebox.showwarning("Already Booked",
                f"MRDNO {p['mrdno']} is already booked on {day}:\n"
                f"{existing['doctor']}  ·  {existing['slot']}")
            return

        taken = is_slot_booked(self.bookings, doc, day, slot)
        if taken:
            messagebox.showwarning("Slot Taken",
                f"{slot} for {doc} on {day} is already booked.\n"
                f"Please select another slot.")
            return

        pred_min = self.patient_data.get("pred_min", 60)
        msg = (f"Confirm Booking\n\n"
               f"MRDNO     : {p['mrdno']}\n"
               f"Patient   : {p['gender']}  ·  {p['age_cat']}\n"
               f"Specialty : {p['specialty']}\n"
               f"Flow      : {p['flow']}\n"
               f"Est. Time : {int(pred_min)} min\n"
               f"Doctor    : {doc}\n"
               f"Day       : {day}\n"
               f"Time Slot : {slot}\n\n"
               f"Confirm?")
        if messagebox.askyesno("Confirm Booking", msg):
            booking = {
                "id":        f"{p['mrdno']}_{day}_{slot}_{doc}".replace(" ","_"),
                "mrdno":     p["mrdno"],
                "gender":    p["gender"],
                "age_cat":   p["age_cat"],
                "specialty": p["specialty"],
                "visit_cat": p["visit_cat"],
                "flow":      p["flow"],
                "doctor":    doc,
                "day":       day,
                "slot":      slot,
                "pred_min":  round(pred_min, 1),
                "booked_at": datetime.now().isoformat(),
            }
            self.bookings.append(booking)
            save_bookings(self.bookings)
            messagebox.showinfo("Booked ✓",
                f"Appointment confirmed!\n{p['mrdno']}  →  {doc}  |  {day}  {slot}")
            self._render_patient()

    # ── Manual entry ──────────────────────────────────────────────────
    def _open_manual(self):
        dlg = tk.Toplevel(self)
        dlg.title("Manual Patient Entry")
        dlg.configure(bg=BG)
        dlg.geometry("420x560")
        dlg.grab_set()
        dlg.resizable(False, False)

        # Dialog header
        hdr = tk.Frame(dlg, bg=CARD2, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="✏   Manual Patient Entry",
                 font=self.fnt_head, bg=CARD2, fg=TEXT).pack(side="left", padx=16)
        tk.Frame(dlg, bg=ACCENT2, height=2).pack(fill="x")

        body = tk.Frame(dlg, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=8)

        def lbl(text):
            tk.Label(body, text=text, font=self.fnt_small,
                     bg=BG, fg=TEXT_DIM).pack(anchor="w", pady=(8,2))

        def combo(var, choices):
            cb = ttk.Combobox(body, textvariable=var, values=choices,
                              font=self.fnt_body, state="readonly")
            cb.pack(fill="x")
            return cb

        def entry(var, readonly=False):
            ent_wrap = tk.Frame(body, bg=BORDER, padx=1, pady=1)
            ent_wrap.pack(fill="x")
            state = "readonly" if readonly else "normal"
            ent = tk.Entry(ent_wrap, textvariable=var, font=self.fnt_body,
                           bg=CARD if not readonly else CARD2,
                           fg=TEXT if not readonly else TEAL,
                           insertbackground=ACCENT,
                           relief="flat", bd=4, state=state)
            ent.pack(fill="x")
            if not readonly:
                ent.bind("<FocusIn>",  lambda e: ent_wrap.config(bg=ACCENT))
                ent.bind("<FocusOut>", lambda e: ent_wrap.config(bg=BORDER))
            return ent

        v_mrdno  = tk.StringVar(master=dlg, value="MANUAL")
        v_gender = tk.StringVar(master=dlg, value="Male")
        v_age    = tk.StringVar(master=dlg, value="31-45")
        v_spec   = tk.StringVar(master=dlg, value="Retina")
        v_flow   = tk.StringVar(master=dlg, value="Non-Dilated")
        v_vcat   = tk.StringVar(master=dlg, value="MRE")
        v_avg    = tk.StringVar(master=dlg, value="—")

        lbl("MRDNO / Label");         entry(v_mrdno)
        lbl("Gender");                combo(v_gender, ["Male","Female"])
        lbl("Age Category");          combo(v_age,    AGE_CATS)
        lbl("Specialty");             cb_spec = combo(v_spec, SPECIALTIES)
        lbl("Patient Flow Type");     cb_flow = combo(v_flow, FLOW_TYPES)
        lbl("Visit Category");        combo(v_vcat,   ["MRE","SRE","REG"])

        # ── Predicted workup row (read-only, auto-filled) ─────────────
        lbl("Predicted Workup (min)")
        pred_wrap = tk.Frame(body, bg=BORDER, padx=1, pady=1)
        pred_wrap.pack(fill="x")
        pred_entry = tk.Entry(pred_wrap, textvariable=v_avg, font=self.fnt_body,
                              bg=CARD2, fg=TEAL, relief="flat", bd=4,
                              state="readonly")
        pred_entry.pack(fill="x", side="left", expand=True)
        pred_lbl = tk.Label(pred_wrap, text="auto", font=self.fnt_small,
                            bg=CARD2, fg=TEXT_DIM, padx=8)
        pred_lbl.pack(side="right")

        def _auto_predict(*_):
            """Called whenever specialty or flow type changes — re-runs model."""
            if self.pkg is None:
                v_avg.set("—")
                pred_lbl.config(text="no model", fg=TEXT_DIM)
                return
            try:
                spec     = v_spec.get()
                flow     = v_flow.get()
                # Use a neutral consultant/session/arrival for manual entry
                predicted = predict_workup(
                    self.pkg, spec, flow,
                    consultant="",          # no consultant → uses global mean
                    session="Forenoon",
                    arrival="09:00 - 10:00",
                    avg_workup=0.0,         # no history → model-only prediction
                )
                v_avg.set(str(int(predicted)))
                pred_lbl.config(text="✓ predicted", fg=TEAL)
            except Exception:
                v_avg.set("—")
                pred_lbl.config(text="error", fg=RED)

        # Bind to both dropdowns so it updates live
        cb_spec.bind("<<ComboboxSelected>>", _auto_predict)
        cb_flow.bind("<<ComboboxSelected>>", _auto_predict)

        # Run once on open so it's not blank
        dlg.after(100, _auto_predict)

        def apply():
            try:
                avg = float(v_avg.get())
            except ValueError:
                avg = 60.0
            self.patient_data = {
                "mrdno":       v_mrdno.get(),
                "gender":      v_gender.get(),
                "age_cat":     v_age.get(),
                "specialty":   v_spec.get(),
                "visit_cat":   v_vcat.get(),
                "flow":        v_flow.get(),
                "consultant":  "",
                "designation": "Specialist",
                "session":     "Forenoon",
                "arrival":     "09:00 - 10:00",
                "past_visits": 1,
                "avg_workup":  avg,
            }
            dlg.destroy()
            self._render_patient()

        btn_frame = tk.Frame(dlg, bg=BG)
        btn_frame.pack(fill="x", padx=20, pady=12)

        apply_btn = tk.Label(btn_frame, text="  Show Schedule →  ",
                              font=self.fnt_body, bg=ACCENT2, fg=TEXT,
                              cursor="hand2", padx=6, pady=7,
                              highlightbackground=ACCENT, highlightthickness=1)
        apply_btn.pack(side="right")
        apply_btn.bind("<Button-1>", lambda e: apply())
        apply_btn.bind("<Enter>", lambda e: apply_btn.config(bg=ACCENT, fg=BG))
        apply_btn.bind("<Leave>", lambda e: apply_btn.config(bg=ACCENT2, fg=TEXT))

        cancel_lbl = tk.Label(btn_frame, text="  Cancel  ", font=self.fnt_small,
                               bg=CARD, fg=TEXT_DIM, cursor="hand2",
                               padx=4, pady=7,
                               highlightbackground=BORDER, highlightthickness=1)
        cancel_lbl.pack(side="right", padx=(0,8))
        cancel_lbl.bind("<Button-1>", lambda e: dlg.destroy())
        cancel_lbl.bind("<Enter>", lambda e: cancel_lbl.config(bg=CARD2, fg=TEXT))
        cancel_lbl.bind("<Leave>", lambda e: cancel_lbl.config(bg=CARD,  fg=TEXT_DIM))


# ── Entry point ────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = OPDSchedulerApp()
    style = ttk.Style(app)
    style.theme_use("clam")
    style.configure("TCombobox",
        fieldbackground=CARD, background=CARD,
        foreground=TEXT, selectbackground=ACCENT2,
        selectforeground=TEXT, arrowcolor=TEXT_DIM,
        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
    style.map("TCombobox",
        fieldbackground=[("readonly", CARD)],
        foreground=[("readonly", TEXT)],
        background=[("readonly", CARD)])
    app.mainloop()