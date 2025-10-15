# STAKE Automated Timetable System

## Project Overview
* The STAKE Automated Timetable System is a Python-based automation tool that generates clash-free academic timetables for IIIT Dharwad.
* It aims to reduce manual scheduling effort, minimize human error, and ensure optimal utilization of classrooms, labs, and faculty resources.
* The system reads input Excel sheets containing course details, faculty names, and L-T-P structures, and produces structured, color-coded timetables automatically.

## Key Features
* Automatic clash-free timetable generation
* Faculty workload balancing and availability checks
* Lecture, Tutorial, and Lab slot scheduling (L-T-P mapping)
* Support for electives and merged divisions
* Unit testing setup for validation of functions
* Excel output with color-coded and merged cells
* (In Progress) Debugging of slot allotment and elective scheduling logic

## Current Status
* Phase	Status	Notes
* Requirements Gathering - Completed	
* System Design	- Completed	
* Coding & Implementation	- In Progress (core logic implemented, under debugging)
* Unit Testing	- Test cases added (validation of helper functions and scheduling)
* Integration & UI	- Planned for next stage
 
## Tech Stack
* Language: Python 3
* Libraries Used:
    pandas — Data processing
    openpyxl — Excel file generation and formatting
    datetime, random, math — Slot management and time calculations
    unittest — Unit testing framework
    Version Control: Git + GitHub

## Repository Structure
STAKE-Timetable-Scheduler/

|
|-- data/                 Department-wise Excel input files

|
|-- tests/                Unit testing and validation

 \-- TestCases.md      Documented test cases and expected outputs

 \-- test_inputs/      Sample Excel inputs for unit testing

|
|-- docs/                 DPR and design-related documents

|
|-- main.py               Core timetable generator (currently under debugging)

|
|-- README.md             Project overview and status

## Team
* Sachin Kumar -	24BCS125
* T Rekha Sree -	24BCS152
* P Haswanth Reddy -	24BCS096
* Sampath S Koralli	- 24BCS129

Guided by: Dr. Vivekraj VK

## Next Steps
* Continue debugging and verifying slot allocation logic
* Ensure tutorial and elective scheduling match L-T-P configuration
* Add constraint validation (break hours, room conflicts, faculty load)
* Generate final formatted Excel timetable output
* Begin UI development phase