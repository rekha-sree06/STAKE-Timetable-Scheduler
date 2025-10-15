||Test case input||Description||Expected output||

||"3-0-0-0-2"||Test `parse_LTP_from_ltpsc`||Returns `(3,0,0)`||

||"Neha Bharti, Rohit Rai"||Test `parse_people`||Returns `['Neha Bharti','Rohit Rai']`||

||div1.csv||Test `read_input_file`||Reads CSV correctly; 16 rows loaded; column names normalized||

||Division `1CSEB` with non-elective courses||Test `build_slot_requests_for_division`||Generates normal slots for non-electives; basket `ELECTIVE-1` for electives; durations split correctly according to `slot_durations`||

||Elective basket `ELECTIVE-1` spanning merged divisions `1CSEB,1DSAI,1ECE`||Test `schedule_globally`||Slot scheduled in all divisions; no faculty/room conflicts; respects breaks||

||Minimum gap = 0, break slots `12:30-13:30` and `16:30-17:00`||Test `schedule_globally`||No slot scheduled in break slots; breaks shaded in Excel||

||Generated Excel output (first_halfsem & second_halfsem)||Test `write_year_excel`||Excel file with merged cells, reference table correct, colored cells per slot, break slots shaded; saved in `timetable_outputs`||

||Unscheduled slots due to conflicts||Test `schedule_globally`||Slot appears in `unscheduled` list with proper message||

||Lec/Tut/Lab hours splitting||Test `build_slot_requests_for_division`||Hours correctly split according to `slot_durations` (lec=1.5h, tut=1h, lab=2h), remainder handled||

||Merged divisions in normal slots||Test `schedule_globally`||Slots assigned to merged divisions properly, respecting faculty and room constraints||

||Reference table in Excel||Test `write_year_excel`||Shows Slot Name, Course Code, Course Title, Faculty, L-T-P-S-C, ROOM.NO, LAB ROOM.NO correctly; one row per slot base||

||Slot colors based on SLOT NAME||Test `write_year_excel`||Cells colored deterministically per `slot_base` for clarity||
