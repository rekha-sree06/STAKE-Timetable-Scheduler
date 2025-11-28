# Development Progress Report (DPR)

##  Project: STAKE – Automated Timetable & Exam Scheduling System  
### Phase: Coding • Debugging • Integration • Documentation

---

## 🛠 Tools & Technologies
- **Language:** Python 3.8+  
- **Libraries:**  
  - `pandas` – Data processing and Excel handling  
  - `openpyxl` – Excel generation, formatting, merged cells  
  - `datetime`, `math`, `random` – Slot and timing calculations  
  - `json` – Settings configuration (`settings.json`)  
  - `unittest` – Testing scripts  
- **Version Control:** Git + GitHub  

---

##  Work Completed (Updated)
### 1. Timetable Generation (`main.py`)
- Implemented robust **academic timetable scheduling** algorithm  
- Features:
  - Lecture / Tutorial / Lab (L-T-P) slot calculation  
  - Merged divisions synchronization (MERGE column)  
  - Support for electives, full-semester and half-semester courses  
  - Break slots and working hours configurable via `settings.json`  
- Implemented constraints:
  - Clash detection: faculty, room, division  
  - Faculty workload balancing  
  - Minimum gap enforcement between consecutive slots  
- Excel output:
  - Division-wise sheets  
  - Minute-level interval headers  
  - Color-coded and merged cells for easy readability  
  - "Reference Table" and "Unallotted Slots" sheet for debugging  
- Integrated division strength and room capacity logic  

### 2. Exam Scheduler & Seating System (`exam.py`)
- Fully automated exam timetable generator  
- Features:
  - Split into **First-Half** and **Second-Half** timetables  
  - Room allotment based on number of students  
  - Seating arrangement generator (per day, per room, FN/AN sessions)  
  - Invigilator allocation with per-invigilator schedule  
- Generated **complete EXAM_OUTPUT** folder including:
  - `firsthalf_timetable.xlsx` and `secondhalf_timetable.xlsx`  
  - Day-wise seating Excel files  
  - Invigilator schedules  
- Input validation:
  - Reads `Rooms.xlsx` and `invigilators_list.xlsx`  
  - Uses exact headers for courses (`COURSE CODE`, `FACULTY`, `L-T-P-S-C`, `NO. OF STUDENTS`)  

### 3. Repository Enhancements
- Structured repository with proper folders:
  - `data/` – Input Excel files  
  - `EXAM_OUTPUT/` – Exam outputs with FIRSTHALF / SECONDHALF  
  - `TT_Output/` – Semester timetable outputs  
  - `docs/` – DPR and documentation  
  - `tests/` – Sample test cases and input files  
- Added input Excel files for all 4 years (CSE, DSAI, ECE)  
- Structured DPR and README documentation for easy onboarding  

---

##  Current Focus (Ongoing Development)
- Fine-tuning **L-T-P block allocation** for edge cases  
- Ensuring elective courses are properly distributed in merged divisions  
- **Room allotment automation** for regular semester timetable  
- Validating constraints:
  - Faculty workload limits  
  - Break hour enforcement  
  - Clash handling across multiple divisions  
- Improving Excel formatting and reference tables for usability  
- Optimizing seating and invigilator allocation logic  

---

##  Current Stage
- **Timetable generator:** Core logic stable; polishing constraints  
- **Exam scheduler:** Fully functional; outputs validated  
- **Seating system:** Reliable with room-capacity logic and daily FN/AN sessions  
- **Documentation:** Updated and structured; includes README.md and DPR.md  
- **Repository:** Organized with clear folder structure and sample inputs  

---

##  Next Steps
- Complete **room allotment automation** for normal semester timetable  
- Implement optional **UI** (Tkinter/Flask) for user-friendly operation  
- Add **ZIP packaging** for downloadable timetable and exam outputs  
- Expand **error reporting & validation** for user inputs  
- Strengthen **unit testing** for edge cases and merged division scenarios  
- Enhance exam seating algorithm for strict load balancing of invigilators  

---

##  Team
- Sachin Kumar – 24BCS125  
- T. Rekha Sree – 24BCS152  
- P. Haswanth Reddy – 24BCS096  
- Sampath S. Koralli – 24BCS129  

**Guide:** Dr. Vivekraj V K
