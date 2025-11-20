# STAKE Automated Timetable & Exam Scheduling System

The **STAKE Automated Timetable System** is a comprehensive Python-based automation tool designed to generate **clash-free academic timetables**, **exam schedules**, **invigilator allocation**, and **seating arrangements** for IIIT Dharwad.  
It significantly reduces manual workload, minimizes scheduling errors, and ensures efficient utilization of faculty, classrooms, and resources.

---

## 📌 Project Overview

This system automatically generates:

- **Academic Timetables** (Years 1–4, First & Second Half Semester)
- **Exam Timetables** (First Half & Second Half)
- **Invigilator Allocation**
- **Seating Arrangements for All Exam Days**

The system reads structured **Excel input files** containing faculty, course, L-T-P, and room details and outputs formatted Excel sheets.

---

## ⭐ Key Features

### 🔹 Academic Timetables
- Fully automated timetable generation.
- Strict conflict checking:
  - Faculty availability
  - Room availability
  - Merged divisions (CSE/DS/AI)
  - L–T–P slot mapping  
- Balances workload across faculty.
- Generates **color-coded Excel timetables** with merged cells.

### 🔹 Exam Scheduling
- Auto-generated exam schedule for first and second half semesters.
- Ensures no clashes across departments and years.
- Balanced subject distribution across exam days.

### 🔹 Invigilator Allocation
- Automatic fair distribution of invigilation duties.
- No faculty overload.
- Priority-based allocation logic.

### 🔹 Seating Arrangement Automation
- Room-capacity based seat allocation.
- Mixed-branch and mixed-year seating support.
- Day-wise Excel output for all exam phases.

### 🔹 Testing & Validation
- Dedicated `tests/` module.
- Validates helper functions and scheduling logic.
- Includes sample inputs and documented test cases.

---

## 📂 Repository Structure

```
timetable-scheduler/
│   exam.py                → Exam timetable, invigilators & seating generator
│   main.py                → Academic timetable generator
│   README.md              → Project documentation
│   requirements.txt       → Dependencies
│
├───data/                  → Input Excel files
│       1CSEA.xlsx
│       1CSEB.xlsx
│       ...
│       invigilators_list.xlsx
│       Rooms.xlsx
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
        Timetable_Year1_first_halfsem.xlsx
        ...
```

---

## 🛠️ Tech Stack

- **Language**: Python 3
- **Libraries**:
  - `pandas` — Data handling  
  - `openpyxl` — Excel writing & formatting  
  - `datetime`, `math`, `random` — Scheduling logic  
  - `unittest` — Testing
- **Version Control**: Git + GitHub

---

## ⚙️ Installation

### 1. Clone the repository
```bash
git clone https://github.com/<your-username>/timetable-scheduler.git
cd timetable-scheduler
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

---

## ▶️ Usage

### 🔸 Generate Academic Timetables
```bash
python main.py
```
Outputs stored in:
```
TT_Output/
```

### 🔸 Generate Exam Timetable, Invigilators & Seating
```bash
python exam.py
```
Outputs stored in:
```
EXAM_OUTPUT/
```

---

## 📄 Input File Requirements

Place all `.xlsx` files inside the `data/` folder.

Required files:
- **Branch & Year Course Files**  
  e.g., `1CSEA.xlsx`, `1DSAI.xlsx`, `3ECE.xlsx`, ...
- **Rooms.xlsx**  
  Room names, capacities
- **invigilators_list.xlsx**  
  Faculty list for allocation

All sheets must follow the exact format of provided sample files.

---

## 🧪 Testing

Test resources are under:
```
tests/
tests/test_inputs/
```

Includes:
- Test cases for helper functions  
- Sample inputs for validation  
- Expected behaviour documentation  

Run tests manually or integrate into future CI workflows.

---

## 📘 Documentation

Full **Detailed Project Report (DPR)** is available in:
```
docs/DPR.md
```

Includes:
- System architecture  
- Flow diagrams  
- Algorithm explanation  
- Constraint logic  
- Implementation details  

---

## 👥 Team

| Name                | Roll Number |
|---------------------|-------------|
| Sachin Kumar        | 24BCS125    |
| T Rekha Sree        | 24BCS152    |
| P Haswanth Reddy    | 24BCS096    |
| Sampath S Koralli   | 24BCS129    |

**Guided by: Dr. Vivekraj VK**

---

## 🚧 Next Steps

- Automate room allotment for regular academic timetables.
- Improve elective and tutorial slot mapping.
- Strengthen conflict validation across merged divisions.
- Enhance faculty load balancing and break-hour constraints.
- UI/GUI development (desktop/web interface).
- Add auto-visualization of timetable.
- Expand unit testing and edge-case handling.

---

## 📜 License

This project is intended for academic timetable automation.  
Feel free to modify and extend for institutional needs.

---

## 🎯 Summary

This system automates the complete timetable and exam scheduling workflow — from lecture scheduling to invigilation and seating. It replaces hours of manual work with a fully consistent, conflict-free scheduler.