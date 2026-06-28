"""
OPD Scheduler — Full Stress Test  (v2)
Simulates one complete Monday with realistic patient mix:
  49.4% MRE  13.1% SRE  37.5% REG  (from EHR n=225,961)
  44.9% Non-Dilated  22.2% Dilated  33% Procedure
Consultation slot = per-patient per-FLOW average from CONS_WORKUP_TIME.
Target: place ~350 patients (hospital daily throughput).
"""

import os, sys, json, logging, random, time
from datetime import datetime
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import pandas as pd

# ── Wipe old files ────────────────────────────────────────────────────
for fname in ["bookings.json","scheduler.log","patient_lookup.csv"]:
    p = os.path.join(BASE, fname)
    if os.path.exists(p): os.remove(p)

logging.basicConfig(
    filename=os.path.join(BASE,"scheduler.log"),
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    force=True,
)
log = logging.getLogger("opd.stresstest")
console = logging.StreamHandler(sys.stdout)
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(levelname)-8s  %(message)s"))
log.addHandler(console)

from scheduling import (
    SCHEDULE, SPECIALTIES, DAYS_ORDER, FLOW_TYPES,
    PERIOD_START_MIN, PERIOD_END_MIN, PERIOD_DUR_MIN,
    WALKIN_ALLOWED_PERIODS, WALKIN_PERIOD_CAP,
    load_patient_lookup, load_model, predict_workup,
    rank_slots_v2, has_interval_overlap, is_already_booked,
    save_bookings, clock_range, total_opd_minutes,
    get_broad_idx_from_minute, GLOBAL_CONSULT_MEDIAN, get_consult_slot,
    BOOKINGS_FILE, LOOKUP_CACHE,
)

# ═══════════════════════════════════════════════════════════════════
# Step 1 — Build lookup
# ═══════════════════════════════════════════════════════════════════
log.info("="*64)
log.info("STRESS TEST v2  START  %s", datetime.now().isoformat())
log.info("="*64)
log.info("Step 1: Building patient_lookup.csv")
t0 = time.time()
lookup = load_patient_lookup()
log.info("  %d unique patients  (%.1fs)", len(lookup), time.time()-t0)

# Verify flow columns exist
flow_cols = [c for c in lookup.columns if c.startswith("avg_consult_")]
log.info("  Flow consult columns: %s", flow_cols)

# ═══════════════════════════════════════════════════════════════════
# Step 2 — Load model
# ═══════════════════════════════════════════════════════════════════
log.info("Step 2: Loading workup model")
try:
    pkg = load_model(); log.info("  OK")
except Exception as e:
    pkg = None; log.warning("  Failed (%s) — demo mode", e)

# ═══════════════════════════════════════════════════════════════════
# Step 3 — Build patient pool (real EHR MRDNOs)
# ═══════════════════════════════════════════════════════════════════
log.info("Step 3: Sampling real patients — target 350-patient day")

# Realistic mix from EHR analysis image
# 62.5% follow-up (MRE+SRE), 37.5% walk-in (REG)
# Flow: 45% Non-Dilated, 22% Dilated, 33% Procedure
TARGET     = 400   # oversample so ~350 actually place after caps
TARGET_DAY = "Mon"
SPEC_CLEAN = [s.strip() for s in SPECIALTIES]

# Strip whitespace in lookup index just in case
lookup.index = lookup.index.astype(str).str.strip()

VCAT_FLOW_QUOTA = [
    # (visit_cat, flow,          n_per_spec)   — 5 specs × these = pool size
    ("MRE", "Non-Dilated",  30),   # 5×30=150
    ("MRE", "Dilated",      10),   # 5×10=50
    ("MRE", "Procedure",    16),   # 5×16=80
    ("SRE", "Non-Dilated",   8),   # 5×8 =40
    ("SRE", "Dilated",       4),   # 5×4 =20
    ("SRE", "Procedure",     5),   # 5×5 =25
    ("REG", "Non-Dilated",  22),   # 5×22=110
    ("REG", "Dilated",       8),   # 5×8 =40
]  # total EHR pool ~515 unique patients

def _sample(spec, vcat, flow, n):
    spec = spec.strip()
    cands = lookup[
        (lookup.get("Specialty", pd.Series(dtype=str)).str.strip() == spec) &
        (lookup.get("Patient visit category", pd.Series(dtype=str)).str.strip() == vcat) &
        (lookup.get("Patient Flow type", pd.Series(dtype=str)).str.strip() == flow)
    ]
    if cands.empty:
        cands = lookup[
            (lookup.get("Specialty", pd.Series(dtype=str)).str.strip() == spec) &
            (lookup.get("Patient visit category", pd.Series(dtype=str)).str.strip() == vcat)
        ]
    if cands.empty: return []
    sampled = cands.sample(min(n, len(cands)), random_state=42)
    out = []
    for mrdno, r in sampled.iterrows():
        slot = get_consult_slot(r, flow)
        out.append({
            "mrdno":       str(mrdno),
            "gender":      "Female" if str(r.get("Gender","")).startswith("F") else "Male",
            "age_cat":     str(r.get("Age category","31-45")).strip(),
            "specialty":   spec,
            "visit_cat":   vcat,
            "flow":        flow,
            "consultant":  str(r.get("Consultant Name","")),
            "session":     str(r.get("Session","Forenoon")),
            "arrival":     str(r.get("Arrival hour","09:00 - 10:00")),
            "avg_workup":  float(r.get("avg_workup", 60)) if pd.notna(r.get("avg_workup")) else 60.0,
            "consult_min": slot,
            "source":      "EHR",
        })
    return out

pool = []
for spec in SPEC_CLEAN:
    for vcat, flow, n in VCAT_FLOW_QUOTA:
        pool.extend(_sample(spec, vcat, flow, n))

log.info("  EHR pool: %d patients", len(pool))

# ═══════════════════════════════════════════════════════════════════
# Step 4 — Manual walk-ins (no EHR history, consult_min from global median)
# ═══════════════════════════════════════════════════════════════════
log.info("Step 4: Manual walk-ins")

MANUAL = [
    ("WALK-001","Male",  "61-75","Retina",                  "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-002","Female","46-60","Glaucoma",                 "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-003","Male",  "31-45","Cornea",                   "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-004","Female","76-90","General",                  "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-005","Male",  "13-19","Pediatric and Low vision", "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-006","Female","20-30","Retina",                   "Dilated",     GLOBAL_CONSULT_MEDIAN["Dilated"]),
    ("WALK-007","Male",  "61-75","Glaucoma",                 "Dilated",     GLOBAL_CONSULT_MEDIAN["Dilated"]),
    ("WALK-008","Female","46-60","Cornea",                   "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-009","Male",  "31-45","General",                  "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-010","Female","0-12", "Pediatric and Low vision", "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-011","Male",  "76-90","Retina",                   "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-012","Female","61-75","Glaucoma",                 "Dilated",     GLOBAL_CONSULT_MEDIAN["Dilated"]),
    ("WALK-013","Male",  "46-60","General",                  "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-014","Female","20-30","Cornea",                   "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-015","Male",  "13-19","Pediatric and Low vision", "Dilated",     GLOBAL_CONSULT_MEDIAN["Dilated"]),
    ("WALK-016","Female","31-45","Retina",                   "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-017","Male",  "46-60","Glaucoma",                 "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-018","Female","61-75","Cornea",                   "Dilated",     GLOBAL_CONSULT_MEDIAN["Dilated"]),
    ("WALK-019","Male",  "20-30","General",                  "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
    ("WALK-020","Female","76-90","Retina",                   "Non-Dilated", GLOBAL_CONSULT_MEDIAN["Non-Dilated"]),
]

for (mrdno,gender,age_cat,spec,flow,slot) in MANUAL:
    pool.append({
        "mrdno":mrdno,"gender":gender,"age_cat":age_cat,
        "specialty":spec,"visit_cat":"REG","flow":flow,
        "consultant":"","session":"Forenoon","arrival":"09:00 - 10:00",
        "avg_workup":60.0,"consult_min":slot,"source":"MANUAL_WALKIN",
    })
log.info("  Manual walk-ins: %d", len(MANUAL))
log.info("  Total pool: %d", len(pool))

# ═══════════════════════════════════════════════════════════════════
# Step 5 — Shuffle realistically and book
# ═══════════════════════════════════════════════════════════════════
log.info("Step 5: Booking on %s", TARGET_DAY)

# Interleave: don't dump all MRE first then REG; arrival is random
random.seed(2024)
random.shuffle(pool)

bookings       = []
placed         = []
rej_dup        = []
rej_cap        = []

spec_stats  = {s: {"placed":0,"rej":0} for s in SPEC_CLEAN}
vcat_placed = defaultdict(int)
flow_placed = defaultdict(int)
manual_placed = []

slot_dist = []   # consult durations of placed patients

t0 = time.time()
for p in pool:
    mrdno = p["mrdno"]; day = TARGET_DAY
    spec  = p["specialty"]; vcat = p["visit_cat"]; flow = p["flow"]
    dur   = p["consult_min"]

    if is_already_booked(bookings, mrdno, day):
        rej_dup.append(mrdno); continue

    ranked = rank_slots_v2(spec, day, flow, bookings, dur, top_n=1, visit_cat=vcat)
    if not ranked:
        rej_cap.append({"mrdno":mrdno,"spec":spec,"vcat":vcat,"flow":flow,"dur":dur})
        if spec in spec_stats: spec_stats[spec]["rej"] += 1
        continue

    best = ranked[0]
    bid  = f"{mrdno}_{day}_{best['start_minute']}_{best['doc']}".replace(" ","_")
    bk   = {
        "id":               bid,
        "mrdno":            mrdno,
        "gender":           p["gender"],
        "age_cat":          p["age_cat"],
        "specialty":        spec,
        "visit_cat":        vcat,
        "flow":             flow,
        "doctor":           best["doc"],
        "day":              day,
        "start_minute":     best["start_minute"],
        "duration_minutes": best["duration_minutes"],
        "end_minute":       best["start_minute"] + best["duration_minutes"],
        "pred_min":         round(p["avg_workup"], 1),
        "consult_min":      dur,
        "source":           p["source"],
        "booked_at":        datetime.now().isoformat(),
        "clock_range":      best["clock_range"],
        "score":            round(best["score"], 4),
    }
    bookings.append(bk)
    placed.append(bk)
    slot_dist.append(dur)

    if spec in spec_stats: spec_stats[spec]["placed"] += 1
    vcat_placed[vcat] += 1
    flow_placed[flow]  += 1
    if p["source"] == "MANUAL_WALKIN": manual_placed.append(bk)

    log.info(
        "BOOKED  %-12s  %-30s  %-4s  %-14s  %s  (%dmin)  sc=%.3f  [%s]",
        mrdno, spec, vcat, flow, best["clock_range"],
        dur, best["score"], p["source"][:3]
    )

elapsed = time.time() - t0

# ═══════════════════════════════════════════════════════════════════
# Step 6 — Save
# ═══════════════════════════════════════════════════════════════════
log.info("Step 6: Saving %d bookings", len(bookings))
save_bookings(bookings)

# ═══════════════════════════════════════════════════════════════════
# Step 7 — Integrity audit
# ═══════════════════════════════════════════════════════════════════
log.info("Step 7: Integrity audit")
by_doc = defaultdict(list)
for b in bookings: by_doc[b["doctor"]].append(b)

overlaps = 0
for doc, bk in by_doc.items():
    bk.sort(key=lambda x: x["start_minute"])
    for j in range(len(bk)-1):
        a, b_ = bk[j], bk[j+1]
        if a["end_minute"] > b_["start_minute"]:
            overlaps += 1
            log.error("OVERLAP  %s  %s %s vs %s %s",
                      doc, a["mrdno"], a["clock_range"], b_["mrdno"], b_["clock_range"])

reg_violations = [b for b in bookings
                  if b["visit_cat"]=="REG"
                  and get_broad_idx_from_minute(b["start_minute"]) not in WALKIN_ALLOWED_PERIODS]
log.info("  Overlaps: %d  REG violations: %d", overlaps, len(reg_violations))

# ═══════════════════════════════════════════════════════════════════
# Step 8 — Capacity report
# ═══════════════════════════════════════════════════════════════════
PERIOD_LABELS = ["08-09","09-10","10-11","11-12","12-13:30",
                 "13:30-14:30","14:30-15:30","15:30-17"]
pfu=[0]*8; preg=[0]*8; pall=[0]*8; pct_fu=[0]*8; pct_reg=[0]*8

for b in bookings:
    bi = get_broad_idx_from_minute(b["start_minute"])
    if 0<=bi<8:
        pall[bi] += b["duration_minutes"]
        (preg if b["visit_cat"]=="REG" else pfu)[bi] += b["duration_minutes"]

total_opd = sum(total_opd_minutes(ds)
                for s in SPEC_CLEAN
                for ds in SCHEDULE.get(s,{}).get(TARGET_DAY,[]))
total_booked = sum(b["duration_minutes"] for b in bookings)

log.info("="*68)
log.info("CAPACITY REPORT — %s", TARGET_DAY)
log.info("%-14s  %6s  %7s  %6s  %5s  %5s  %s",
         "Period","FU-min","REG-min","Total","FU%","REG%","Status")
log.info("-"*68)
for i, lbl in enumerate(PERIOD_LABELS):
    cap_i = sum(PERIOD_DUR_MIN[i] for s in SPEC_CLEAN
                for ds in SCHEDULE.get(s,{}).get(TARGET_DAY,[])
                if i<len(ds) and ds[i]=="OPD")
    if cap_i == 0: continue
    fp = round(pfu[i]/cap_i*100,1); rp = round(preg[i]/cap_i*100,1)
    ap = round(pall[i]/cap_i*100,1)
    status = "FULL" if pall[i]>=cap_i else ("NEAR FULL" if ap>85 else "")
    if i not in WALKIN_ALLOWED_PERIODS and preg[i]>0: status="REG LEAK"
    log.info("%-14s  %6d  %7d  %6d  %4.1f%%  %4.1f%%  %s",
             lbl, pfu[i], preg[i], pall[i], fp, rp, status)
log.info("-"*68)
log.info("OPD capacity : %d min  |  Booked: %d min  (%.1f%%)",
         total_opd, total_booked, total_booked/total_opd*100)
log.info("="*68)

# ═══════════════════════════════════════════════════════════════════
# Step 9 — Summary
# ═══════════════════════════════════════════════════════════════════
avg_slot = sum(slot_dist)/len(slot_dist) if slot_dist else 0
log.info("SUMMARY")
log.info("  Pool attempted   : %d", len(pool))
log.info("  PLACED           : %d  (EHR=%d  MANUAL=%d)",
         len(placed),
         sum(1 for b in placed if b["source"]=="EHR"),
         len(manual_placed))
log.info("  Rejected (dup)   : %d", len(rej_dup))
log.info("  Rejected (cap)   : %d", len(rej_cap))
log.info("  Avg consult slot : %.1f min  (range %d-%d)",
         avg_slot, min(slot_dist), max(slot_dist))
log.info("  Booking speed    : %.2fs for %d calls", elapsed, len(pool))
log.info("  Overlaps         : %d", overlaps)
log.info("  REG violations   : %d", len(reg_violations))
log.info("")
log.info("  By visit category:")
for vc in ["MRE","SRE","REG"]:
    log.info("    %-4s : %d placed", vc, vcat_placed[vc])
log.info("  By flow:")
for fl in FLOW_TYPES:
    log.info("    %-14s : %d placed", fl, flow_placed[fl])
log.info("  By specialty:")
for sp in SPEC_CLEAN:
    log.info("    %-30s : %d placed  %d rejected",
             sp, spec_stats[sp]["placed"], spec_stats[sp]["rej"])
log.info("")
log.info("  Manual walk-ins placed (%d/%d):", len(manual_placed), len(MANUAL))
for b in manual_placed:
    log.info("    %-12s  %-30s  %-14s  %s  (%dmin)",
             b["mrdno"], b["specialty"], b["flow"], b["clock_range"], b["duration_minutes"])

if rej_cap:
    log.info("")
    log.info("  Capacity-rejected patients:")
    for r in rej_cap:
        log.info("    %-12s  %-30s  %-4s  %-14s  dur=%dmin",
                 r["mrdno"], r["spec"], r["vcat"], r["flow"], r["dur"])

log.info("="*68)

print(f"\nFiles: {BOOKINGS_FILE}  |  scheduler.log  |  {LOOKUP_CACHE}")
print(f"Result: {len(placed)} placed / {len(pool)} attempted  |"
      f"  overlaps={overlaps}  reg_violations={len(reg_violations)}"
      f"  util={total_booked/total_opd*100:.1f}%")
