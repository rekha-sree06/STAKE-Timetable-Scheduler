# STAKE Timetable & Exam Scheduler  
Automated Timetable Generator & Seating + Invigilation System  
Developed by **Team STAKE**

---

##  Overview

**STAKE-Timetable-Scheduler** is an end-to-end automation system that generates:

- Class Timetables (Year-wise & Sem-wise)  
- Exam Seating Plans  
- Invigilation Schedules  
- Consolidated Excel Outputs

The system processes input Excel files and produces **conflict-free, faculty-balanced, room-optimized** schedules.

This project uses:

- **Python 3.8+**  
- **pandas** — for Excel processing  
- **openpyxl** — for Excel writing  
- **json** — for configurable settings  
- Python standard libraries: `os`, `math`, `random`, `collections`, `pathlib`

---

##  Features

###  Timetable Generation (`main.py`)
- Reads division-wise files (e.g., 1CSEA.xlsx, 1CSEB.xlsx)  
- Slot allocation based on L–T–P values (supports `L-T-P-S-C` or `L-T-P`)  
- Ensures:
  - No faculty clashes  
  - No room overlaps  
  - Synchronized merged divisions (via `MERGE` column)  
  - Proper L/T/P block scheduling  
  - Minimum gap between consecutive slots (configurable)  
  - Faculty gap enforcement (configurable)  
- Writes per-year, per-half Excel timetables

###  Exam Scheduling (`exam.py` / `seating_scheduler_final_seating_sessions.py`)
- Groups courses into exam slots
- Allocates sessions (FN/AN) based on room capacities
- Creates per-day seating grids per room (rows × columns)
- Assigns invigilators and generates per-invigilator schedules
- Outputs per-half directories with seating files and invigilation schedule

---

##  Project Folder Structure

    timetable-scheduler/
    │   exam.py                → Exam timetable, invigilators & seating generator
    │   main.py                → Academic timetable generator
    │   README.md              → Project documentation
    │   requirements.txt       → Dependencies
    │
    ├───data/                  → Input Excel files
    │   ├── Rooms.xlsx
    │   ├── invigilators_list.xlsx
    │   ├── 1CSEA.xlsx
    │   ├── 1CSEB.xlsx
    │   ├── 1DSAI.xlsx
    │   ├── 1ECE.xlsx
    │   ├── 2CSEA.xlsx
    │   ├── 2CSEB.xlsx
    │   ├── 2DSAI.xlsx
    │   ├── 2ECE.xlsx
    │   ├── 3CSEA.xlsx
    │   ├── 3CSEB.xlsx
    │   ├── 3DSAI.xlsx
    │   ├── 3ECE.xlsx
    │   ├── 4CSEA.xlsx
    │   ├── 4CSEB.xlsx
    │   ├── 4DSAI.xlsx
    │   ├── 4ECE.xlsx
    │
    ├───docs/
    │       DPR.md             → Full project report
    │
    ├───EXAM_OUTPUT/
    │   ├───FIRSTHALF/
    │   │       firsthalf_timetable.xlsx
    │   │       Invigilator_Schedules.xlsx
    │   │       seating_arrangements/Day_1.xlsx ...
    │   │
    │   └───SECONDHALF/
    │           secondhalf_timetable.xlsx
    │           Invigilator_Schedules.xlsx
    │           seating_arrangements/Day_1.xlsx ...
    │
    ├───tests/
    │       TestCases.md
    │       test_inputs/
    │           1CSEAI.xlsx
    │
    └───TT_Output/
    │   ├── Timetable_Year1_second_halfsem.xlsx
    │   ├── Timetable_Year2_first_halfsem.xlsx
    │   ├── Timetable_Year2_second_halfsem.xlsx
    │   ├── Timetable_Year3_first_halfsem.xlsx
    │   ├── Timetable_Year3_second_halfsem.xlsx
    │   ├── Timetable_Year4_first_halfsem.xlsx
    │   ├── Timetable_Year4_second_halfsem.xlsx

---

## 📦 Requirements

Create a `requirements.txt` with the following (or run the command below):

Installation command:

    pip install pandas openpyxl

(These are the only non-standard libraries required; others are from Python stdlib.)

---

## ⚙️ settings.json (example)

Place `settings.json` at repo root to override defaults. Example content:

    {
      "working_days": ["Mon", "Tue", "Wed", "Thu", "Fri"],
      "working_hours": ["09:00", "18:30"],
      "break_slots": ["12:30-13:30", "16:30-17:00"],
      "slot_durations": {"lec": 1.5, "lab": 2.0, "tut": 1.0}
    }

If `settings.json` is missing, `main.py` uses sensible defaults shown above.

---

## ▶️ How to run — Timetable Generator (`main.py`)

1. Ensure input division files (e.g., `1CSEA.xlsx`) are placed in `data/` or accessible paths.
2. Run:

    python main.py

3. Interactive prompts (you will be asked):
   - Minimum gap between consecutive slots in minutes (default 5)
   - Minimum faculty gap (default 180 minutes)
   - Number of academic years
   - For each year: number of divisions, division short-name and path to file

4. Outputs:
   - `TT_Output/Year_<Y>/Timetable_Year<Y>_first_halfsem.xlsx`
   - `TT_Output/Year_<Y>/Timetable_Year<Y>_second_halfsem.xlsx`

Each workbook contains:
- Division-wise sheets with minute-accurate interval headers
- A "Reference Table" with input course rows
- "Unallotted Slots" sheet for items that couldn't be scheduled

---

## ▶️ How to run — Exam Scheduler (`exam.py`)

> Note: The exam script in your repo (`seating_scheduler_final_seating_sessions.py`) contains a hardcoded `divisions` dictionary using `project\...` paths. Either update those paths to point to `data\...` OR place files accordingly.

1. Edit `seating_scheduler_final_seating_sessions.py` (top) to adjust file paths if required:

    - `divisions = { ... }` (map years & division names to file paths)
    - `rooms_path = r"data\Rooms.xlsx"`
    - `invig_path = r"data\invigilators_list.xlsx"`

2. Run:

    python seating_scheduler_final_seating_sessions.py
    # or
    python exam.py  (if you renamed the file back to exam.py)

3. Outputs (per half):

    EXAM_OUTPUT/
      FIRSTHALF/
        firsthalf_timetable.xlsx
        Invigilator_Schedules.xlsx
        seating_arrangements/
          Day_1.xlsx
          Day_2.xlsx
          ...
      SECONDHALF/
        secondhalf_timetable.xlsx
        Invigilator_Schedules.xlsx
        seating_arrangements/
          Day_1.xlsx
          ...

Each Day_N.xlsx contains `FN` and `AN` sheets with room grids and a `REFERENCE` sheet mapping slots to sessions.

---

## Input Excel file requirements (exact headings used by the code)

### Division files (e.g., `1CSEA.xlsx`)
Each row = a course offering. Required/used columns (exact strings preferred):

- `ELECTIVE OR NOT` (YES / NO)
- `FULLSEM OR HALFSEM` (e.g., FULLSEM, HALFSEM-1, HALFSEM-2)
- `COURSE CODE`
- `COURSE TITLE`
- `FACULTY` (comma-separated allowed)
- `CLASS ASSISTANTS` (optional)
- `LAB ASSISTANTS` (optional)
- `L-T-P-S-C` **or** `L-T-P` (e.g., `3-1-0`)
- `ROOM.NO` (comma-separated allowed)
- `LAB ROOM.NO` (comma-separated allowed)
- `SLOT NAME`
- `MERGE` (comma-separated; merged marker used to sync slots)
- `NO. OF STUDENTS` (integer, required for exam scheduling)

### Rooms file (`Rooms.xlsx`)
Columns required:
- `Room`
- `Seating Capacity`

### Invigilators file (`invigilators_list.xlsx`)
At least one column (`NUMBER`). Preferably two columns (`NUMBER`, `NAME`). Script uses first two columns.

---

## Common issues & troubleshooting

- **File not found**: Provide correct paths when `main.py` prompts or edit `divisions` dict in exam script.
- **Slots unplaced**: Inspect `Unallotted Slots` sheet. Common causes: wrong L-T-P format, insufficient rooms, conflicting MERGE entries, or faculty availability constraints.
- **Exam capacity insufficient**: Increase room list or change session capacity policy (the code currently uses `cap // 2` as usable seats per session).
- **Invigilator distribution**: Current logic is best-effort; modify `allocate_seating_for_session()` if you need strict/load-balanced rules.

---

## Suggested quick edits

- To point exam script to `data/` files, update top of `seating_scheduler_final_seating_sessions.py`:

    # Example (edit the paths)
    divisions = {
      1: {"1CSEA": r"data\1CSEA.xlsx", ...},
      ...
    }
    rooms_path = r"data\Rooms.xlsx"
    invig_path = r"data\invigilators_list.xlsx"

- To use full room capacity instead of half (session policy), replace occurrences of `cap // 2` with `cap`.

- To use roll numbers in final seating labels, add provision in `allocate_seating_for_session()` to read rolllists per division and pop roll numbers instead of generating numeric suffixes.

---

## Authors & Acknowledgements

Developed by **Team STAKE**:
- Sachin Kumar (24BCS125)
- T. Rekha Sree (24BCS152)
- P. Haswanth Reddy (24BCS096)
- Sampath S. Koralli (24BCS129)

Guided by: Dr. Vivekraj V K

---

Which would you like next?
