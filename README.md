## `ui_for_scheduling/` — OPD Workup Scheduler

A Tkinter GUI for scheduling OPD consultations. Looks up patients by MR number, predicts their workup time, and books them into 15-minute consultation slots across a consultant schedule.

### Files

| File | Description |
|------|-------------|
| `scheduling.py` | Main GUI application |
| `bookings.json` | Persisted booking records (auto-updated by the app) |
| `patient_lookup.csv` | Auto-built patient lookup cache derived from EHR data |
| `Consultant schedule.docx` | Source consultant availability schedule |
| `ehr_age7_ar1.xlsx` | Dataset (copy) |
| `workup_model_age7_ar1.pkl` | Model (copy) |

### Features
- Patient lookup by MR number with auto-populated demographic fields
- ML-based workup time prediction at booking time
- 15-minute slot grid (08:00–17:00) per consultant per day
- Bookings persisted to `bookings.json`
- Consultant schedule parsed from `Consultant schedule.docx`

### How to run
```bash
cd ui_for_scheduling
python scheduling.py
```

---
