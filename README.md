# OPD Workup Scheduler — Real-Time Monitoring UI

> Production-grade ophthalmology outpatient scheduling engine with variable-duration bookings, split two-slot dilation workflow, Plan B priority windows, XGBoost workup prediction, and a dark-themed real-time monitoring UI built entirely in Python/Tkinter.

---

## Table of Contents

1. [What this system does](#1-what-this-system-does)
2. [Background and motivation](#2-background-and-motivation)
3. [Architecture overview](#3-architecture-overview)
4. [Files and directory structure](#4-files-and-directory-structure)
5. [EHR data and the patient lookup](#5-ehr-data-and-the-patient-lookup)
6. [The three patient types](#6-the-three-patient-types)
7. [The three flow types](#7-the-three-flow-types)
8. [The dilated two-slot workflow](#8-the-dilated-two-slot-workflow)
9. [Plan B scheduling windows](#9-plan-b-scheduling-windows)
10. [Consultation slot sizing — the core insight](#10-consultation-slot-sizing--the-core-insight)
11. [How the slot-finding engine works](#11-how-the-slot-finding-engine-works)
12. [How slots are scored and ranked](#12-how-slots-are-scored-and-ranked)
13. [Doctor naming and specialty isolation](#13-doctor-naming-and-specialty-isolation)
14. [Capacity guard logic](#14-capacity-guard-logic)
15. [The doctor schedule grid](#15-the-doctor-schedule-grid)
16. [XGBoost workup time prediction](#16-xgboost-workup-time-prediction)
17. [Booking lifecycle](#17-booking-lifecycle)
18. [Legacy booking migration](#18-legacy-booking-migration)
19. [The UI — page by page](#19-the-ui--page-by-page)
20. [Real-time features](#20-real-time-features)
21. [Keyboard shortcuts](#21-keyboard-shortcuts)
22. [Stress test](#22-stress-test)
23. [Key bugs found and fixed](#23-key-bugs-found-and-fixed)
24. [Constants quick reference](#24-constants-quick-reference)
25. [Running the system](#25-running-the-system)

---

## 1. What this system does

This system schedules consultation appointments for patients arriving at an ophthalmology OPD (Outpatient Department). It:

- Looks up a patient by their Medical Record Number (MRDNO) from a 225,961-visit EHR dataset
- Predicts how long their total hospital stay will be (XGBoost model)
- Computes the actual doctor **consultation slot** from their personal EHR history of sign-in/sign-out times — not the total stay
- For Dilated patients, books **two linked slots** — a short initial check, then a post-dilation fundus exam — with the 28-minute dilation wait happening in the waiting area so the doctor is free to see other patients
- Finds the best available gap in the relevant doctor's schedule using Plan B priority windows
- Books with zero overlap, full audit logging, and a live visual timeline

The system is designed to place **~400+ patients per day** across **5 specialties** and **17 doctors** with genuine variable-duration appointments, not fixed 15-minute blocks.

---

## 2. Background and motivation

### Why not fixed slots?

Most hospital schedulers give everyone a fixed 10 or 15 minute slot. In ophthalmology this fails badly:

- A Non-Dilated follow-up is reviewed and done in **2–4 minutes**
- A Dilated patient's *doctor face time* is only **4 + 10 = 14 minutes** across two brief interactions — but naively blocking 38 minutes wasted the entire dilation wait as idle doctor time
- A Procedure patient coming for a pre-op OPD check needs **12 minutes** (the procedure itself is in the OT, not counted here)

Forcing all three into 15-minute blocks wastes 70% of capacity for Non-Dilated patients and wastes 24 minutes of doctor time per Dilated patient during the dilation wait.

### Why separate consultation time from total workup time?

The XGBoost model predicts **total time the patient stays in the hospital** — from registration through workup, dilation wait, consultation, pharmacy, and discharge. That number (average 36–90 minutes depending on flow) is **not** what the doctor needs a slot for. The doctor only needs the patient in the consultation room for the CONS_WORKUP_TIME: the face-to-face time between `CONS_SIGN_IN` and `CONS_SIGN_OUT` in the EHR.

These two numbers are deliberately kept separate in this system:
- `pred_min` — XGBoost predicted total stay (shown informational, not used for booking)
- `consult_min` — per-patient per-flow average of `CONS_WORKUP_TIME` (used to size the booked slot)

### What is Plan A vs Plan B? (Lin, Jin & Chia 2014)

The DMAIC scheduling framework (Lin, Jin & Chia 2014) distinguishes:

**Plan A** — Unstructured first-come-first-served. All patients arrive and queue without any assigned time window. Afternoon patients pile up because mornings fill by walk-in. Doctor idle time and patient waiting time are both high.

**Plan B** — Structured time windows by patient type and workup requirements. Each patient type (follow-up vs walk-in) and each flow (Procedure, Dilated, Non-Dilated) gets assigned windows that match the clinical logic:
- Procedures need the OT-adjacent morning block (8–10am) so they transition to OT on time
- Dilated patients need the 8am–12pm block because the dilation wait must complete during the OPD session, not after — and the two-slot split means both the check and the exam land within this window
- Non-Dilated follow-ups are most flexible (10am–3pm)
- Walk-in new patients get two dedicated windows (9–11am and 2:30–5pm) so they don't displace follow-ups

This system implements Plan B fully. Every booking decision checks which Plan B window the proposed slot falls in and penalises slots outside the preferred window in the scoring function.

### Why utilization went down when throughput went up

After implementing the split dilation workflow, the utilization metric (booked minutes / total OPD minutes) dropped from 86.2% to 76.4% even though 74 more patients were placed. This is not a regression — it is the metric becoming more accurate.

The old 38-minute dilated slot counted 28 minutes of idle doctor time (dilation wait) as "booked". Now that time is correctly freed for other patients. Those filler patients are counted in the patient total but their short slots (4–12 min) don't restore the inflated minute count the dilation wait artificially provided.

**Patients/hour** is the correct throughput metric for this workflow, not utilization percentage. It is now shown on the dashboard.

---

## 3. Architecture overview

```
scheduling.py
├── Constants & palette
├── SCHEDULE dict             — per-specialty, per-day, per-doctor grid (OPD/OT/LUNCH/NA)
├── load_patient_lookup()     — builds/reads patient_lookup.csv from EHR Excel
├── get_consult_slot()        — returns per-patient per-flow consultation duration (capped)
├── predict_workup()          — XGBoost-based total stay prediction
├── find_gaps()               — scans OPD periods for free time of required duration
├── find_dilation_pairs()     — finds (check_start, exam_start) slot pairs for Dilated patients
├── score_slot_v2()           — multi-factor slot scorer
├── rank_slots_v2()           — gap finder + scorer + capacity guard, returns top N
├── OPDSchedulerApp           — Tkinter GUI
│   ├── Dashboard page        — KPIs, patients/hour, specialty bars, heatmap, recent bookings
│   ├── Scheduler page        — patient card, flow cards, timeline, recommendation banner
│   └── Appointments page     — filterable table, paired cancel, CSV export
└── ToastManager              — non-blocking notification overlay

stress_test.py
├── Wipes existing files
├── Builds lookup
├── Samples real EHR patients by specialty × visit_cat × flow
├── Adds 20 manual walk-ins
├── Shuffles and books one full Monday (handles dilated pairs)
└── Integrity audit + capacity report + summary log
```

---

## 4. Files and directory structure

| File | Purpose |
|------|---------|
| `scheduling.py` | Main application — all logic + UI |
| `stress_test.py` | Headless full-day simulation |
| `workup_model_age7_ar1.pkl` | Trained XGBoost model package |
| `ehr_age7_ar1.xlsx` | Source EHR data (225,961 visits, 134,659 patients) |
| `patient_lookup.csv` | Auto-built cache: one row per MRDNO with averaged fields |
| `bookings.json` | Persistent booking store — Dilated patients have two records each |
| `scheduler.log` | Full audit log (every booking, rejection, overlap check) |

`patient_lookup.csv` and `bookings.json` are generated automatically and do not need to exist before first run.

---

## 5. EHR data and the patient lookup

The EHR Excel file has 225,961 visit rows across 134,659 unique patients. Key columns used:

| EHR column | Usage |
|-----------|-------|
| `MRDNO` | Patient identifier (index key) |
| `VISITDATE` | Used to pick the most recent visit's metadata |
| `Patient Flow type` | `Non-Dilated` / `Dilated` / `Procedure` |
| `Patient visit category` | `MRE` / `SRE` / `REG` |
| `Specialty` | Which department |
| `TOTAL_WORKUP_TIME` | Total hospital stay duration (time object → minutes) |
| `CONS_WORKUP_TIME` | Doctor face time only: `CONS_SIGN_OUT − CONS_SIGN_IN` |
| `Consultant Name` | For workup prediction encoding |
| `Gender`, `Age category` | For patient card display |
| `Session`, `Arrival hour` | For workup prediction |

`load_patient_lookup()` builds `patient_lookup.csv` in one pass:

1. Strip whitespace from all categorical columns (EHR has trailing spaces on flow type values — without this, `"Non-Dilated "` doesn't match `"Non-Dilated"` in groupby and everything becomes NaN)
2. Convert `TOTAL_WORKUP_TIME` and `CONS_WORKUP_TIME` time objects to integer minutes
3. Set zero or negative `consult_min` values to `None` (bad EHR entries)
4. For each MRDNO, take the most recent visit for demographic/specialty info
5. Compute `past_visits` (count), `avg_workup` (mean total stay)
6. Compute **three separate consultation averages** — one per flow type:
   - `avg_consult_nondilated`
   - `avg_consult_dilated`
   - `avg_consult_procedure`

The three-way split is critical. Without it, a patient who mostly had Procedure visits (74-min consults due to full procedure suite stay being recorded) would get a 74-minute OPD slot for a Non-Dilated appointment — inflating their slot by 18× and choking the schedule.

> **Note on Dilated EHR averages**: `avg_consult_dilated` in the EHR reflects the old single-slot workflow where the doctor held the patient for the entire 38-minute dilation visit. Under the new two-slot workflow this value is irrelevant — the initial check is always capped at 6 minutes via `CONSULT_SLOT_CAP["Dilated"] = 6`, discarding the historical average entirely.

---

## 6. The three patient types

| Code | Name | Description | Priority |
|------|------|-------------|----------|
| `MRE` | Major Return | Existing follow-up patient, scheduled recall | Highest — gets first access to all OPD periods |
| `SRE` | Short Return | Existing follow-up, shorter recall interval | High — same windows as MRE |
| `REG` | Registration | New walk-in patient, no prior appointment | Restricted — two windows only, 37.5% cap |

**Patient mix from EHR**: 49.4% MRE · 13.1% SRE · 37.5% REG

Follow-up patients (MRE + SRE = 62.5%) are explicitly prioritised over walk-ins. The system guarantees this by hard-locking REG to specific periods and enforcing a per-period minute cap.

---

## 7. The three flow types

| Flow | Booking model | Doctor time |
|------|--------------|-------------|
| `Non-Dilated` | Single slot, 2–30 min (median 4 min) | Quick review, no dilation. Many patients done in 2 minutes. |
| `Dilated` | **Two linked slots** — 4-min check + 10-min exam, separated by 28-min dilation wait | Doctor sees other patients during the wait. Total doctor time: 14 min. Old single-slot model wasted 24 min per patient. |
| `Procedure` | Single slot, 1–15 min (median 12 min) | OPD slot is the pre-procedure check only. Patient is booked into an OPD period exactly like any other patient. The actual procedure happens separately in OT — that booking is outside this system entirely. `CONS_WORKUP_TIME` in EHR was recorded as the full procedure suite stay (~76 min), so it is hard-capped at 15 min. |

**Flow mix from EHR**: 44.9% Non-Dilated · 22.2% Dilated · 33% Procedure

---

## 8. The dilated two-slot workflow

This is the most significant architectural feature. A dilated patient requires two doctor interactions separated by a waiting period for the dilation drops to take effect biologically.

### Old workflow (single slot, wasteful)

```
08:00  Patient enters → doctor checks, instils drops
08:04  Doctor waiting (patient sitting in room while drops work)  ← 28 min idle
08:32  Doctor does fundus exam
08:38  Patient leaves
```
Slot held: 08:00–08:38 = **38 min**, of which 28 min the doctor was idle.

### New workflow (two slots, efficient)

```
08:00  Patient enters → doctor checks, instils drops       [SLOT 1: 4 min]
08:04  Patient moves to waiting area                       [doctor free]
08:04  Doctor books other patients during this window      [28 min free]
08:32  Patient recalled → doctor does fundus exam          [SLOT 2: 10 min]
08:42  Patient leaves
```
Doctor time consumed: **14 min** across two brief slots. 24 min freed for other patients.

### Implementation

Two separate booking records are created in `bookings.json`, linked by a shared `dilation_pair_id`:

```json
{
  "id": "123456_Mon_480_RET-Doc1",
  "mrdno": "123456",
  "flow": "Dilated",
  "doctor": "RET-Doc1",
  "start_minute": 480,
  "duration_minutes": 4,
  "end_minute": 484,
  "dilation_pair_id": "DIL_123456_Mon_480_RET-Doc1",
  "dilation_phase": "check"
}

{
  "id": "123456_Mon_480_RET-Doc1_exam",
  "mrdno": "123456",
  "flow": "Dilated",
  "doctor": "RET-Doc1",
  "start_minute": 512,
  "duration_minutes": 10,
  "end_minute": 522,
  "dilation_pair_id": "DIL_123456_Mon_480_RET-Doc1",
  "dilation_phase": "exam"
}
```

`exam_start_minute ≥ check_end_minute + DILATION_WAIT_MIN` is enforced by `find_dilation_pairs()`.

### Constants

```python
DILATION_WAIT_MIN = 28   # biological minimum for drops to dilate pupils fully
DILATION_EXAM_MIN = 10   # post-dilation fundus exam duration
GLOBAL_CONSULT_MEDIAN["Dilated"] = 4   # initial check duration
CONSULT_SLOT_CAP["Dilated"]      = 6   # max initial check, clamps old EHR 38-min avg
```

### How `find_dilation_pairs` works

```python
def find_dilation_pairs(slot1_periods, slot2_periods, sorted_bookings,
                        initial_dur, wait_min, exam_dur):
```

1. Calls `find_gaps(slot1_periods, ...)` to get all candidate check start times
2. For each candidate `t1`:
   - Computes `earliest_t2 = t1 + initial_dur + wait_min`, snapped to SCAN_STEP
   - Adds `t1` as a temporary occupied block so exam search avoids it
   - Scans `slot2_periods` starting from `earliest_t2` for the first free 10-min gap
   - If found, records the pair `(t1, t2)` and moves on
3. Returns list of `(check_start, exam_start)` tuples

Note: `slot1_periods` respects walk-in period restrictions (REG can only check in allowed periods). `slot2_periods` uses all OPD periods — the patient is already inside the clinic during the wait, so no arrival window restriction applies to the exam slot.

### Cancellation

Cancelling either record in an Appointments page row cancels **both** records in the pair using the shared `dilation_pair_id`. The UI confirms "Cancel BOTH dilation slots?" before deleting.

### REG period restriction for exam slots

For REG Dilated patients, the initial check (slot 1) is restricted to `WALKIN_ALLOWED_PERIODS = {1, 2, 6, 7}`. The exam slot (slot 2) is exempt from this restriction — it is the continuation of a booking already in progress, not a new walk-in arrival. The stress test integrity audit skips `dilation_phase == "exam"` records when checking REG violations.

---

## 9. Plan B scheduling windows

Defined in `PLAN_B_WINDOWS` as a dict mapping `(visit_cat, flow)` → list of period indices.

The 8:00–17:00 day is divided into 8 broad periods:

| Index | Label | Start | End |
|-------|-------|-------|-----|
| 0 | Early | 08:00 | 09:00 |
| 1 | Early | 09:00 | 10:00 |
| 2 | Mid | 10:00 | 11:00 |
| 3 | Mid | 11:00 | 12:00 |
| 4 | Mid | 12:00 | 13:30 |
| 5 | Lunch | 13:30 | 14:30 |
| 6 | Afternoon | 14:30 | 15:30 |
| 7 | Afternoon | 15:30 | 17:00 |

```
(MRE, Procedure)     → [0, 1]           8–10am (early: must start before OT takes over)
(SRE, Procedure)     → [0, 1]           8–10am
(MRE, Dilated)       → [0, 1, 2, 3]    8am–12pm (both check and exam land in this window)
(SRE, Dilated)       → [0, 1, 2, 3]    8am–12pm
(MRE, Non-Dilated)   → [2, 3, 4, 6, 7] 10am–3pm flexible
(SRE, Non-Dilated)   → [2, 3, 4, 6, 7] 10am–3pm flexible
(REG, *)             → [1, 2, 6, 7]    9–11am + 2:30–5pm only
```

Slots outside the preferred window are penalised in the score (+0.6 penalty = heavily discouraged but not impossible if no preferred slots remain).

---

## 10. Consultation slot sizing — the core insight

```python
GLOBAL_CONSULT_MEDIAN = {
    "Non-Dilated": 4,
    "Dilated":     4,   # initial check only; exam is separately DILATION_EXAM_MIN=10
    "Procedure":   12,
}

CONSULT_SLOT_CAP = {
    "Non-Dilated": 30,
    "Dilated":     6,   # clamps old EHR 38-min avg — irrelevant under new split workflow
    "Procedure":   15,
}
```

`get_consult_slot(patient_row, flow)` returns the slot in this order of preference:

1. Patient's personal average `CONS_WORKUP_TIME` for **this specific flow type** (e.g. `avg_consult_nondilated` if today is Non-Dilated)
2. Clamp to `CONSULT_SLOT_CAP[flow]` — prevents one outlier patient from monopolising the schedule, and for Dilated specifically ensures the old 38-min EHR avg never gets used as a slot size
3. If no personal history exists for this flow: fall back to `GLOBAL_CONSULT_MEDIAN[flow]`

**This function must always be used for slot sizing** — computing `consult_min` from `raw_consult` directly without going through `get_consult_slot` will bypass the cap and produce 44-minute initial checks for Dilated patients whose EHR average was recorded under the old single-slot workflow.

There is **no hardcoded 15-minute fallback** for all patients. Some Non-Dilated patients have a genuine 2-minute average and that is what gets booked.

---

## 11. How the slot-finding engine works

### Standard single slot — `find_gaps`

`find_gaps(opd_periods, sorted_bookings, duration_needed)` takes a list of `(start_min, end_min)` tuples representing OPD periods for a specific doctor and returns a list of candidate start minutes.

```
For each OPD period (ps → pe):
  t = ps
  While t + duration_needed ≤ pe:
    Check if any existing booking overlaps [t, t+duration_needed)
    If overlap:
      Jump t forward to end of that booking, align to next SCAN_STEP multiple
      Retry
    If no overlap:
      Check OVERLAP_BUFFER: does the next booking start within duration_needed + 3 minutes?
      If yes: t += SCAN_STEP, continue (need breathing room before next patient)
      If no:  add t to candidates, t += SCAN_STEP
```

**`SCAN_STEP = 5`** — candidates are always at 5-minute aligned boundaries. This prevents the schedule from fragmenting into odd-minute slots that are hard to communicate to patients.

**`OVERLAP_BUFFER = 3`** — 3 minutes of minimum gap between consecutive appointments. Accounts for patient transition time and prevents marginal timing errors.

Interval arithmetic is half-open: `[start, end)`. A booking from 09:00 to 09:15 occupies minutes 540–555. The next booking starting at 09:15 (555) does **not** overlap.

### Dilated two-slot — `find_dilation_pairs`

`find_dilation_pairs(slot1_periods, slot2_periods, sorted_bookings, initial_dur, wait_min, exam_dur)` extends the gap-finding logic to return pairs:

```
For each slot1 candidate t1 from find_gaps(slot1_periods, ...):
  earliest_t2 = t1 + initial_dur + wait_min  (snapped to SCAN_STEP)
  temp_bookings = existing_bookings + [synthetic block for slot1]
  For each OPD period that extends past earliest_t2:
    Scan from max(period_start, earliest_t2) for first free exam_dur gap
    If found: record pair (t1, t2), break
```

The synthetic block for slot1 prevents the exam search from accidentally placing slot2 overlapping slot1 in edge cases where the wait is very short.

---

## 12. How slots are scored and ranked

`score_slot_v2(start_min, duration_min, doc_name, day, flow, bookings, visit_cat, doc_slots)` returns a float score. **Lower is better.**

| Component | Weight | Logic |
|-----------|--------|-------|
| Plan B window | 0.40 | 0.0 if in preferred period, 0.6 if outside |
| Doctor load | 0.25 | `booked_minutes / opd_capacity` — prefer less loaded doctors |
| Congestion | 0.15 | `period_total_booked / period_total_opd` — prefer emptier periods |
| Dilated penalty | 0.10 | +0.4 if Dilated appointment is in the last 30% of the day (patient won't complete dilation in time) |
| Lunch penalty | 0.05 | +0.3 if slot is in the 13:30–14:30 lunch period |
| Walk-in crowding | 0.05 | +0.2 if the period is already >80% follow-up (protect new-patient access) |

For Dilated patients, scoring is applied to the **check slot (slot 1)** only. The exam slot follows deterministically from the check slot and is not independently scored.

`rank_slots_v2(spec, day, flow, bookings, pred_duration, top_n, visit_cat)` collects all candidates from all doctors in the specialty, scores each, sorts ascending, returns top N.

The top recommendation is shown in the UI banner and in the `⚡ Recommend` dialog. Rank 1 = gold star, Rank 2 = silver star, Rank 3 = bronze star on the timeline.

---

## 13. Doctor naming and specialty isolation

Each specialty has its own team of doctors who only treat patients from that specialty. This is modelled by naming doctors with a specialty prefix:

```python
spec_prefix = spec.strip().split()[0][:3].upper()
doc_name = f"{spec_prefix}-Doc{di+1}"
# Retina            → RET-Doc1, RET-Doc2, ...
# Glaucoma          → GLA-Doc1, GLA-Doc2, ...
# Cornea            → COR-Doc1, ...
# General           → GEN-Doc1, ...
# Pediatric and Low vision → PED-Doc1, ...
```

This naming is used everywhere: in `has_interval_overlap`, `rank_slots_v2`, the timeline rendering, and `bookings.json`. Without the prefix, "Doc 1" in Retina and "Doc 1" in Glaucoma would be treated as the same physical doctor, causing `has_interval_overlap` to reject cross-specialty bookings that happen to be at the same time.

The `_timeline_view` method computes the same prefix from the specialty name before rendering per-doctor rows, ensuring the doctor labels in the timeline match the `doctor` field stored in bookings.

---

## 14. Capacity guard logic

### REG hard period lock

Walk-in patients (`visit_cat == "REG"`) are only allowed to search for slots in periods `{1, 2, 6, 7}` (9–11am and 2:30–5pm). This is a hard constraint, not a preference:

```python
WALKIN_ALLOWED_PERIODS = {1, 2, 6, 7}

if is_walkin:
    search_periods = [(ps, pe) for i, (ps, pe) in enumerate(...)
                      if i in WALKIN_ALLOWED_PERIODS
                      and doc_slots[i] == "OPD"]
```

If no slot exists in those periods, the patient is rejected (capacity-rejected, not duplicate-rejected).

### Per-period REG minute cap

Even within allowed periods, REG patients cannot consume more than 37.5% of total OPD minutes:

```python
WALKIN_PERIOD_CAP = 0.375

if is_walkin and period_total_opd[broad] > 0:
    if (period_booked_reg[broad] + duration) / period_total_opd[broad] > WALKIN_PERIOD_CAP:
        continue  # skip this candidate
```

37.5% mirrors the actual patient mix from the EHR. This ensures follow-up patients always have access to morning and lunchtime periods even if many walk-ins arrive early.

### Cross-specialty consistency fix

Both `period_total_opd` (the denominator) and `period_booked_reg` (the numerator) must be computed across **all specialties**, not just the current one. If `period_total_opd` is computed for only Glaucoma's 2 doctors (120 min in 09–10), but `period_booked_reg` counts REG minutes from all 17 doctors (67 min), the ratio becomes 67/120 = 55.8% which wrongly blocks Glaucoma follow-ups from the 09–10 period even though the true cross-specialty ratio is only 67/660 = 10.2%.

```python
# CORRECT: all specialties
for _spec in SCHEDULE:
    for _ds in SCHEDULE[_spec].get(day, []):
        for i, s in enumerate(_ds):
            if s == "OPD": period_total_opd[i] += PERIOD_DUR_MIN[i]
```

### Follow-up protection from REG overflow

For follow-up (MRE/SRE) patients, the search skips any period where REG has already consumed more than `WALKIN_PERIOD_CAP` of total OPD minutes. This is a one-directional protection: if walk-ins arrive first and fill their 37.5% quota, follow-up patients cannot be displaced from the remaining 62.5%.

---

## 15. The doctor schedule grid

`SCHEDULE` is a nested dict: `SCHEDULE[specialty][day]` returns a list of lists, where each inner list is one doctor's day broken into the 8 broad periods.

Each period slot is one of:
- `"OPD"` — doctor is available for patient appointments
- `"OT"` — doctor is in the operating theatre (no OPD bookings possible)
- `"LUNCH"` — lunch break
- `"NA"` — not available / no session

Example for Glaucoma Monday:
```python
"Mon": [
    ["OPD","OPD","OPD","OPD","NA","LUNCH","NA","OPD"],   # GLA-Doc1: 330 min OPD
    ["OPD","OPD","OPD","OPD","OPD","LUNCH","OPD","OPD"], # GLA-Doc2: 480 min OPD
    ["OT","OT","OT","OT","OPD","LUNCH","OPD","NA"],      # GLA-Doc3: 150 min OPD (OT morning)
]
```

GLA-Doc3 is in the operating theatre from 8–12, comes out for OPD in the 12–13:30 and 14:30–15:30 slots. `find_gaps` and `find_dilation_pairs` only ever scan periods where `doc_slots[i] == "OPD"` — OT, LUNCH, and NA periods are structurally excluded from all booking searches.

Total OPD capacity on Monday across all 17 doctors: **4,860 minutes**.

---

## 16. XGBoost workup time prediction

`load_model()` uses a `SafeUnpickler` that intercepts `xgboost` class references during unpickling and replaces them with a dummy class. This allows the encoding maps (target-encoded specialty, flow, consultant means) to be loaded even if the runtime XGBoost version differs from the training version.

`predict_workup(pkg, spec, flow, consultant, session, arrival, avg_workup)` blends:

```python
# Base: weighted mean of target-encoded features
base = spec_mean*0.3 + flow_mean*0.5 + consultant_mean*0.2

# Session adjustment
if session == "Afternoon": base *= 0.95  # afternoon sessions slightly faster

# Arrival adjustment
if arrival_hour >= 14: base *= 1.05   # late arrivals spend longer
elif arrival_hour <= 9: base *= 0.95  # early arrivals move faster

# Blend with patient's personal history
predicted = base*0.6 + avg_workup*0.4
```

This is used only for the **informational "expected hospital stay"** display. It has no effect on the actual consultation slot booked.

---

## 17. Booking lifecycle

### Non-Dilated / Procedure (single slot)

1. User types MRDNO → `_lookup()` reads from `patient_lookup.csv`
2. Patient card renders with 3 flow type cards showing predicted total stays
3. `consult_min = get_consult_slot(patient_row, flow)` — always goes through the cap
4. `rank_slots_v2()` returns top 3 candidate slots
5. UI shows recommendation banner and highlights slots on the Gantt timeline
6. User clicks "Book this →" → `_book_slot(doc, start_min, duration_min)` called
7. `has_interval_overlap()` checked before commit
8. One booking record appended to `self.bookings`, saved to `bookings.json`, logged
9. Toast notification shown, timeline redraws

### Dilated (two linked slots)

Steps 1–5 identical. In step 4, `rank_slots_v2` calls `find_dilation_pairs` and returns candidates with extra fields: `exam_start_minute`, `exam_duration_minutes`, `exam_clock_range`, `is_dilation_pair: True`.

The recommendation banner shows: `Check 08:00–08:06 (6 min) → wait 28 min → Exam 08:35–08:45 (10 min)`

6. User clicks "Book this →" → `_book_slot(doc, start_min, duration_min, exam_start, exam_dur)` called
7. Confirmation dialog shows both slots
8. **Two** booking records created with shared `dilation_pair_id`, both saved atomically
9. Toast shows both time ranges

### Booking data schema

Non-Dilated / Procedure record:
```json
{
  "id":               "MRDNO_Day_StartMin_Doctor",
  "mrdno":            "123456",
  "gender":           "Female",
  "age_cat":          "46-60",
  "specialty":        "Retina",
  "visit_cat":        "MRE",
  "flow":             "Non-Dilated",
  "doctor":           "RET-Doc3",
  "day":              "Mon",
  "start_minute":     540,
  "duration_minutes": 4,
  "end_minute":       544,
  "pred_min":         72.4,
  "booked_at":        "2026-06-29T09:15:32.441"
}
```

Dilated check record:
```json
{
  "id":                 "MRDNO_Day_StartMin_Doctor",
  "mrdno":              "123456",
  "flow":               "Dilated",
  "doctor":             "RET-Doc1",
  "start_minute":       480,
  "duration_minutes":   6,
  "end_minute":         486,
  "dilation_pair_id":   "DIL_123456_Mon_480_RET-Doc1",
  "dilation_phase":     "check"
}
```

Dilated exam record:
```json
{
  "id":                 "MRDNO_Day_StartMin_Doctor_exam",
  "mrdno":              "123456",
  "flow":               "Dilated",
  "doctor":             "RET-Doc1",
  "start_minute":       514,
  "duration_minutes":   10,
  "end_minute":         524,
  "dilation_pair_id":   "DIL_123456_Mon_480_RET-Doc1",
  "dilation_phase":     "exam"
}
```

`start_minute` and `end_minute` are minutes from midnight. 480 = 08:00, 514 = 08:34.

---

## 18. Legacy booking migration

If `bookings.json` contains entries from an older version of the system that used a `"slot"` label field (e.g. `"slot": "09:00–09:15"`) instead of `start_minute`/`duration_minutes`, `migrate_booking()` parses the old label and injects the new fields automatically on load. This ensures backward compatibility without any manual file editing.

---

## 19. The UI — page by page

### Sidebar (always visible)

- Hospital logo, app version
- Live clock (seconds resolution) and date — updated every 1 second
- Navigation: Dashboard · Scheduler · Appointments
- Keyboard shortcut reference
- Status indicator: pulsing yellow "Loading" during data load → solid green "Ready" after

### Dashboard page

**KPI row** (4 cards):
- **Patients today** — unique MRDNOs booked today (Dilated patients count as 1, not 2)
- **Patients/hour** — unique patients ÷ 9 OPD hours. This is the primary throughput metric, not utilization percentage, because the two-slot dilation workflow makes utilization (booked minutes / capacity) understate actual doctor busyness
- **Utilization** — booked minutes / total OPD capacity minutes. Lower than the patient count suggests under the new workflow because the 28-min dilation wait is no longer counted as booked time
- **All patients** — unique MRDNOs across all days

**Specialty utilisation bars** — horizontal progress bars per specialty for the current day, colour-coded green/yellow/red.

**Weekly heatmap** — 5 rows (specialties) × 6 columns (days) grid with heat interpolation from green (0%) to red (100%) based on booked minutes.

**Flow breakdown** — 3 cards showing count and percentage of Non-Dilated / Dilated / Procedure bookings today. Exam-phase records are excluded so Dilated shows patient count, not slot count.

**Recent bookings feed** — last 12 bookings in reverse chronological order with MRDNO, specialty, flow, doctor, day, time window, and booking timestamp.

Auto-refreshes every 30 seconds. Timestamp shown in header.

### Scheduler page

**Toolbar**:
- MRDNO entry field with Enter key binding → auto look-up
- Day selector (Mon–Sat)
- Look up / Manual Entry / ⚡ Recommend / 🗑 Clear buttons

**Patient card** — avatar (M/F coloured), MRDNO, specialty, gender, age, past visit count, average workup. Visit category badge (MRE/SRE/REG).

**Already-booked banner** — shown if MRDNO is already booked on the selected day. Links to Appointments page.

**Flow type cards** — 3 clickable cards showing predicted total hospital stay for Non-Dilated / Dilated / Procedure. Selected flow is highlighted. Clicking a card changes the active flow and re-runs slot ranking.

**Info row** — side-by-side: expected hospital stay (informational) vs consultation slot booked (the actual appointment duration from `get_consult_slot`).

**Recommendation banner**:
- Non-Dilated / Procedure: `⚡ Recommended: RET-Doc2 · 09:15–09:19 (4 min)`
- Dilated: `⚡ Recommended: RET-Doc2 · Check 08:00–08:06 (6 min) → wait 28 min → Exam 08:35–08:45 (10 min)`
- Green if in preferred Plan B window, red if outside. Shows doctor load and score.

**Timeline** — one row per doctor showing 8:00–17:00. Segments coloured by period type (OPD/OT/Lunch/NA) with zone tinting (Early/Mid/Afternoon). Existing bookings shown as dark green filled rectangles with duration label. Top-3 recommended slots highlighted with rank stars. Red dashed "now" line during operating hours. Click a recommended slot to book directly.

**Legend** — period type and zone colour key at the bottom.

### Appointments page

Filterable, scrollable table of all bookings:
- Live search across MRDNO, specialty, flow, doctor (Ctrl+F focuses this)
- Day and specialty dropdown filters
- Row hover highlight
- **Cancel button** — for standard bookings: cancels that record. For dilated pair bookings: shows "Cancel BOTH dilation slots?" and removes both records using the shared `dilation_pair_id`
- Export CSV button (Ctrl+E) — prompts for save path, exports all visible fields

---

## 20. Real-time features

| Feature | Interval | Mechanism |
|---------|----------|-----------|
| Live clock | 1 second | `self.after(1000, self._tick_clock)` |
| Statusbar counters | 5 seconds | `self.after(5000, self._tick_statusbar)` (shows unique patient count) |
| Auto-refresh | 30 seconds | `self.after(30_000, self._auto_refresh)` |
| "Now" line on timeline | On render | `datetime.now().hour*60 + datetime.now().minute` |
| LIVE indicator in section header | On render | Checks if current time is within 08:00–17:00 and today matches selected day |
| Loading pulse animation | 400ms steps | `_start_loading_pulse()` / `_start_dot_animation()` |
| Hover animations on nav buttons | 16ms steps | `_animate_hover_in()` / `_animate_hover_out()` with `_lerp_color()` |
| Toast notifications | Auto-dismiss 3.5s | `ToastManager` with queued display |

Data is loaded on a background daemon thread so the UI never freezes on startup.

---

## 21. Keyboard shortcuts

| Shortcut | Action |
|----------|--------|
| `F5` | Refresh current view |
| `Ctrl+F` | Focus search (switches to Appointments if needed) |
| `Ctrl+E` | Export bookings to CSV |
| `Esc` | Clear search field |
| `Enter` (in MRDNO field) | Trigger patient look-up |

---

## 22. Stress test

`stress_test.py` simulates one complete Monday from scratch. Run it to regenerate clean `bookings.json`, `scheduler.log`, and `patient_lookup.csv`.

### What it does

**Step 1** — Wipe existing files. Rebuild `patient_lookup.csv` from EHR Excel.

**Step 2** — Load XGBoost model (falls back to demo mode if unavailable).

**Step 3** — Sample real EHR patients by specialty × visit_cat × flow using `VCAT_FLOW_QUOTA`:

```python
VCAT_FLOW_QUOTA = [
    ("MRE", "Non-Dilated",  30),  # 5 specialties × 30 = 150 patients
    ("MRE", "Dilated",      10),  # 50
    ("MRE", "Procedure",    16),  # 80
    ("SRE", "Non-Dilated",   8),  # 40
    ("SRE", "Dilated",       4),  # 20
    ("SRE", "Procedure",     5),  # 25
    ("REG", "Non-Dilated",  22),  # 110
    ("REG", "Dilated",       8),  # 40
]  # EHR pool ≈ 515 patients
```

**Step 4** — Add 20 manual walk-in patients (`WALK-001` through `WALK-020`) covering all 5 specialties and all flow types. These have no EHR history so their consultation slot is taken from `GLOBAL_CONSULT_MEDIAN`.

**Step 5** — Shuffle the combined pool (`random.seed(2024)` for reproducibility), book one by one:
- Standard patients: call `rank_slots_v2`, create one booking record
- Dilated patients: call `rank_slots_v2`, detect `is_dilation_pair: True`, create **two** linked booking records (check + exam) with shared `dilation_pair_id`

**Step 6** — Save to `bookings.json`.

**Step 7** — Integrity audit:
- Overlap check: all doctors' bookings sorted by `start_minute`, verify no consecutive pair has `a.end_minute > b.start_minute`
- REG violation check: REG patients whose check slot (`dilation_phase != "exam"`) falls outside `WALKIN_ALLOWED_PERIODS`. Exam slots are intentionally exempt.

**Step 8** — Per-period capacity report showing FU minutes, REG minutes, total, and utilisation percentage.

**Step 9** — Summary: placed/attempted, by visit category, by flow, by specialty, manual walk-ins placed, capacity-rejected list.

### Latest results

```
Pool: 535 patients attempted
Placed: 415  (398 EHR + 17 manual walk-ins)
Booking records: 501  (415 standard + 86 dilated exam slots)
Utilisation: 3,715 / 4,860 min = 76.4%
Patients/hour: 46.1
Overlaps: 0    REG violations: 0

By flow:
  Non-Dilated : 254 patients
  Dilated     : 86 patients  (172 booking records)
  Procedure   : 75 patients
```

> **Why 415 > 341 (previous single-slot result)**: The old 38-min dilated slot was blocking 28 min of idle doctor time per dilated patient. The split workflow freed that time, allowing 74 more patients to be placed. Utilization appears lower because the 28-min dilation wait is no longer counted in booked minutes — it was never real doctor work.

---

## 23. Key bugs found and fixed

### Mixed flow consultation average inflating slots

**Bug**: `avg_consult` was a single average over all visits regardless of flow type. A Procedure patient's lifetime average was 79 minutes (full procedure suite stay). Booking them for a Non-Dilated visit gave them a 79-minute OPD slot.

**Fix**: Compute three separate averages — `avg_consult_nondilated`, `avg_consult_dilated`, `avg_consult_procedure` — and pick the one matching today's flow with `get_consult_slot(patient_row, flow)`.

### Procedure slot cap

**Bug**: `CONS_WORKUP_TIME` for Procedure patients covers the entire procedure suite experience (eye drops administration + procedure + recovery room), not just the OPD pre-check. Values of 60–180 minutes appeared in the data and were being used as OPD slot sizes.

**Fix**: `CONSULT_SLOT_CAP["Procedure"] = 15` and `GLOBAL_CONSULT_MEDIAN["Procedure"] = 12`.

### Dilated slot cap not applied in UI

**Bug**: `_render_patient` computed `consult_min` manually as `max(1, round(raw_consult))` without going through `get_consult_slot`. For Dilated patients the EHR average was 38–44 minutes (old single-slot workflow). This bypassed `CONSULT_SLOT_CAP["Dilated"] = 6` and booked a 44-minute initial check instead of a 6-minute one.

**Fix**: Replace the inline computation with `consult_min = get_consult_slot(p, p["flow"])`. All slot sizing must go through this function to guarantee the cap is applied.

### Doctor name collision

**Bug**: All specialties used `f"Doc {di+1}"`. "Doc 1" in Retina and "Doc 1" in Glaucoma resolved to the same string. `has_interval_overlap` then rejected any Glaucoma booking that overlapped in time with a Retina booking, and the timeline rendered the wrong bookings per doctor row.

**Fix**: `doc_name = f"{spec_prefix}-Doc{di+1}"`. Doctors became `RET-Doc1`, `GLA-Doc1`, `PED-Doc1`, etc. The same prefix logic must be applied in both `rank_slots_v2` and `_timeline_view`. Throughput jumped from 252 to 341 patients.

### Timeline doc_name mismatch

**Bug**: `_timeline_view` was constructing `doc_name = f"Doc {di+1}"` (old format without prefix), so `doc_bk` was always empty — no existing bookings were shown on the timeline, and the rank_map check `c["doc"] != doc_name` always failed so no recommended slot highlights appeared.

**Fix**: Add `spec_prefix_tl = spec.strip().split()[0][:3].upper()` before the doctor loop and use `f"{spec_prefix_tl}-Doc{di+1}"`.

### Cross-specialty cap ratio mismatch

**Bug**: `period_total_opd` (denominator for the REG per-period cap check) was computed using only the current specialty's doctors. But `period_booked_reg` (numerator) counted REG minutes across all specialties. For Glaucoma's period 09-10: 2 doctors × 60 min = 120 min denominator, but 67 min REG from all 17 doctors in numerator = 55.8% — wrongly exceeds 37.5% cap and blocked all Glaucoma follow-ups from 09–10 even though the hospital-wide REG consumption was only 10.2%.

**Fix**: Changed `period_total_opd` to loop over `SCHEDULE` for all specialties.

### EHR trailing whitespace

**Bug**: `"Patient Flow type"` column had values like `"Non-Dilated "` with a trailing space. `groupby("Patient Flow type")` created separate groups for `"Non-Dilated"` and `"Non-Dilated "`, resulting in NaN flow averages for most patients.

**Fix**: `.str.strip()` on all three categorical columns at the start of `load_patient_lookup()`.

---

## 24. Constants quick reference

```python
# Timing
SCAN_STEP      = 5    # min: candidates always at 5-min boundaries
OVERLAP_BUFFER = 3    # min: minimum gap between consecutive bookings
DAY_START_MIN  = 480  # 08:00
DAY_END_MIN    = 1020 # 17:00

# Consultation slot sizes
GLOBAL_CONSULT_MEDIAN = {
    "Non-Dilated": 4,
    "Dilated":     4,   # initial check only (Phase 1)
    "Procedure":   12,
}

# Hard caps — no single patient can exceed these regardless of personal EHR average
CONSULT_SLOT_CAP = {
    "Non-Dilated": 30,
    "Dilated":     6,   # clamps old 38-min EHR avg under previous single-slot workflow
    "Procedure":   15,
}

# Dilation two-slot parameters
DILATION_WAIT_MIN = 28   # biological minimum for drops to dilate pupils
DILATION_EXAM_MIN = 10   # post-dilation fundus exam duration

# REG walk-in restrictions
WALKIN_ALLOWED_PERIODS = {1, 2, 6, 7}  # 9-11am and 2:30-5pm only (check slot; exam exempt)
WALKIN_PERIOD_CAP      = 0.375         # 37.5% of any period's OPD minutes

# Scoring weights (lower score = better slot)
# Plan B window:     0.40
# Doctor load:       0.25
# Congestion:        0.15
# Dilated penalty:   0.10 (late-day check slot gets +0.4 if exam won't complete in time)
# Lunch penalty:     0.05 (lunch period gets +0.3)
# Walk-in crowding:  0.05 (+0.2 if period >80% follow-up)

# Specialties and doctors
SPECIALTIES = ["Retina","Glaucoma","Cornea","General","Pediatric and Low vision"]
# Retina: 5 doctors, Glaucoma: 3, Cornea: 3, General: 4, Pediatric: 2 = 17 total
# Total Monday OPD capacity: 4,860 minutes across all 17 doctors
```

---

## 25. Running the system

### Prerequisites

```bash
pip install pandas numpy openpyxl xgboost
```

Python 3.9+ with Tkinter (included in standard Windows Python install).

Required files in the same directory as `scheduling.py`:
- `workup_model_age7_ar1.pkl`
- `ehr_age7_ar1.xlsx`

`patient_lookup.csv` is auto-built on first run from the Excel file and cached for subsequent launches. Building it from scratch takes 30–60 seconds for 225,961 rows.

### Launch the UI

```bash
python scheduling.py
```

The sidebar status dot turns green when data finishes loading. Until then, look-ups return a "still loading" warning.

### Run the stress test

```bash
python stress_test.py
```

This **deletes** `bookings.json`, `scheduler.log`, and `patient_lookup.csv` before running. Do not run it against a live booking file you want to keep.

Output files after the test:
- `bookings.json` — 501 records for Monday (415 patients; 86 dilated patients have 2 records each)
- `scheduler.log` — full audit with every BOOKED line, capacity report, and summary
- `patient_lookup.csv` — rebuilt lookup cache (134,659 patients)

### Demo mode

If `workup_model_age7_ar1.pkl` is missing, the system enters demo mode — the sidebar shows a yellow "Demo mode" indicator. Scheduling still works using the patient's `avg_workup` from the lookup instead of the XGBoost prediction. The consultation slot (`consult_min`) is unaffected since it comes from the EHR averages, not the model.

---

*Built on EHR data: n=225,961 visits · 134,659 unique patients · Plan B framework (Lin, Jin & Chia 2014)*
