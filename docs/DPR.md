# Development Progress Report (DPR)

##  Project: STAKE – Automated Timetable & Exam Scheduling System  
### Phase: Coding • Debugging • Integration

---

## 🛠 Tools & Technologies
- **Language:** Python 3  
- **Libraries:**  
  - `pandas` – Data processing  
  - `openpyxl` – Excel generation & formatting  
  - `datetime`, `math`, `random` – Slot/Time logic  
  - `json` – Settings configuration  
  - `unittest` – Testing  
- **Version Control:** Git + GitHub  

---

##  Work Completed (Updated)
###  1. Timetable Generation (main.py)
- Implemented core timetable scheduling algorithm  
- Added support for:
  - Lecture / Tutorial / Lab L-T-P based slot generation  
  - Merged divisions  
  - Electives and half-semester/ full-semester subjects  
  - Break slots and working hours from `settings.json`  
- Implemented:
  - Clash detection (faculty, room, division)  
  - Load balancing & random distribution  
  - Color-coded Excel export with merged cells  
- Debugged multiple issues related to slot allotment and merged batch handling  
- Added room capacity & division strength integration (under refinement)

###  2. Exam Scheduler & Seating System (exam.py)
- Added fully automated exam scheduling  
- Implemented:
  - First-half / second-half exam timetable split  
  - Room allotment based on student strength  
  - Seating arrangement generator (room-wise Excel)  
  - Daily seating Excel exports  
  - Invigilator allocation with separate schedule export  
- Generated **complete EXAM_OUTPUT** folder with all outputs:
  - First half timetable  
  - Second half timetable  
  - Day-wise seating arrangements  
  - Invigilator schedules  

###  3. Repository Enhancements
- Organized repository with proper structure (`data`, `EXAM_OUTPUT`, `TT_Output`, `tests`, `docs`)  
- Added input Excel files for all years (1–4 CSE, DSAI, ECE)  
- Structured DPR and documentation under `docs/`  
- Added sample test cases in `tests/TestCases.md`  

---

##  Current Focus (Ongoing Development)
- Fixing remaining L-T-P mismatches in complex cases  
- Improving elective distribution across merged divisions  
- Refining **room allotment automation** for regular timetable generation  
- Validating:
  - Faculty workload limits  
  - Break hour enforcement  
  - Clash scenarios under different configurations  
- Improving formatting and reference tables in Excel output  

---

##  Current Stage
The project is in the **debugging, validation, and integration** phase.

- **Timetable generator:** Core logic works; fine-tuning and constraint polishing underway.  
- **Exam scheduler:** Fully functional and producing complete outputs.  
- **Seating system:** Working reliably with room-capacity logic.  
- **Documentation:** Updated and structured; repo ready for next development phase.

---

##  Next Steps
- Complete automation of **room allotment for normal semester timetable**  
- Add UI (Tkinter/Flask) for user-friendly operation  
- Add downloadable ZIP output packaging feature  
- Expand error reporting & validation messages  
- Strengthen unit testing for edge cases  

---

##  Team
- Sachin Kumar – 24BCS125  
- T Rekha Sree – 24BCS152  
- P Haswanth Reddy – 24BCS096  
- Sampath S Koralli – 24BCS129  

**Guide:** Dr. Vivekraj VK
