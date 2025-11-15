# seating_scheduler_final_seating_first.py
import os
import math
import random
from collections import defaultdict
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side

random.seed(42)

# -------------------------
# Helpers
# -------------------------
def safe_int(v):
    try:
        s = str(v).strip()
        if s.lower() in ("", "nan", "none"):
            return 0
        return int(float(s))
    except:
        return 0

def prompt_nonempty(prompt_text):
    v = input(prompt_text).strip()
    while not v:
        v = input(prompt_text).strip()
    return v

def base_slotname(slot_name):
    """Extract base slot name: remove division suffixes and elective _Y suffix."""
    s = str(slot_name).strip().upper()
    if "_Y" in s:
        return s.split("_Y")[0]
    if "_" in s:
        return s.split("_")[0]
    return s

# -------------------------
# Inputs
# -------------------------
def get_user_inputs():
    print("=== Exam Scheduler Inputs ===")
    num_years = int(prompt_nonempty("Number of academic years: "))
    divisions = {}
    for y in range(1, num_years + 1):
        num_divs = int(prompt_nonempty(f"Year {y}, number of divisions: "))
        divisions[y] = {}
        for d in range(1, num_divs + 1):
            sn = prompt_nonempty(f"  Short name for Division {d}: ").upper()
            path = prompt_nonempty(f"     Path to Excel/CSV for {sn}: ")
            if not os.path.exists(path):
                raise FileNotFoundError(f"File not found: {path}")
            divisions[y][sn] = path

    rooms_path = prompt_nonempty("\nPath to rooms Excel (Room, Seating Capacity): ")
    if not os.path.exists(rooms_path):
        raise FileNotFoundError(rooms_path)

    invig_path = prompt_nonempty("Path to faculty/assistants Excel: ")
    if not os.path.exists(invig_path):
        raise FileNotFoundError(invig_path)

    num_days = int(prompt_nonempty("\nNumber of exam days: "))
    sessions_per_day = []
    for i in range(num_days):
        sess_type = prompt_nonempty(f"  Day {i+1}, session(s)? Enter FN, AN, or B: ").upper()
        if sess_type not in ("FN", "AN", "B"):
            raise ValueError("Enter FN, AN, or B")
        sessions_per_day.append({"FN": sess_type in ("FN", "B"), "AN": sess_type in ("AN", "B")})

    return divisions, rooms_path, invig_path, num_days, sessions_per_day, num_years

# -------------------------
# Load courses
# -------------------------
def load_courses(divisions, num_years):
    rows = []
    for year in range(1, num_years + 1):
        for div, path in divisions[year].items():
            df = pd.read_excel(path, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            for _, r in df.iterrows():
                elective = str(r.get("ELECTIVE OR NOT", "")).strip().upper()
                half_type = str(r.get("FULLSEM OR HALFSEM", "")).strip().upper()
                course_code = str(r.get("COURSE CODE", "")).strip()
                course_title = str(r.get("COURSE TITLE", "")).strip()
                slot_raw = str(r.get("SLOT NAME", "")).strip().upper()
                merge_raw = str(r.get("MERGE", "")).strip().upper()
                no_students = safe_int(r.get("NO. OF STUDENTS", 0))

                merge_list = [m.strip().upper() for m in merge_raw.split(",") if str(m).strip()]
                if div not in merge_list:
                    merge_list.append(div)

                slot_name = slot_raw
                if elective == "YES":
                    slot_name = f"{slot_raw}_Y{year}"

                rows.append({
                    "YEAR": year,
                    "DIVISION": div,
                    "ELECTIVE": elective,
                    "FULLSEM_TYPE": half_type,
                    "SLOT": slot_name,
                    "SLOT_RAW": slot_raw,
                    "COURSE_CODE": course_code,
                    "COURSE_TITLE": course_title,
                    "MERGE": merge_list,
                    "NO_STUDENTS": no_students
                })
    df_all = pd.DataFrame(rows)
    for c in ["DIVISION", "SLOT", "SLOT_RAW"]:
        if c in df_all.columns:
            df_all[c] = df_all[c].astype(str).str.upper()
    return df_all

# -------------------------
# Split halves
# -------------------------
def split_half(df):
    first = df[df["FULLSEM_TYPE"].isin(["FULLSEM", "HALFSEM-1"])].copy().reset_index(drop=True)
    second = df[df["FULLSEM_TYPE"].isin(["FULLSEM", "HALFSEM-2"])].copy().reset_index(drop=True)
    return first, second

# -------------------------
# Seating-driven allocation to sessions
# -------------------------
def allocate_slots_by_seating_capacity(courses_df, num_days, sessions_per_day, rooms_df):
    """
    Greedy algorithm:
    - Build unique slots (keyed by SLOT)
      * merged_flag True -> students = max(NO_OF_STUDENTS)
      * merged_flag False -> students = sum(NO_OF_STUDENTS)
      * also keep divisions set for clash prevention
    - Iterate sessions in chronological order (day1 FN, day1 AN, day2 FN, ...)
    - For each session compute total usable seats (sum rooms // 2)
    - For each unassigned slot (in some deterministic order), if:
        - slot.students <= total_usable_seats_remaining AND
        - none of slot.divisions already have an assignment in this (day,session)
      then assign whole slot to this session (do not partially seat).
    - Otherwise defer slot to later session.
    - Continue until all slots are assigned or sessions exhausted (remaining slots will remain unassigned — you can inspect them).
    """
    # Build slots
    slot_map = {}
    for _, r in courses_df.iterrows():
        key = r["SLOT"]
        if key not in slot_map:
            slot_map[key] = {"slot_key": key, "slot_raw": r.get("SLOT_RAW", key), "courses": [], "divisions": set(), "merged_flag": False}
        slot_map[key]["courses"].append(r)
        slot_map[key]["divisions"].add(r["DIVISION"])
        # detect merge: if MERGE length > 1 then it's merged
        if isinstance(r["MERGE"], (list, tuple)) and len(r["MERGE"]) > 1:
            slot_map[key]["merged_flag"] = True

    slots = []
    for k, v in slot_map.items():
        if v["merged_flag"]:
            students = max(int(rr["NO_STUDENTS"]) for rr in v["courses"])
        else:
            students = sum(int(rr["NO_STUDENTS"]) for rr in v["courses"])
        slots.append({
            "slot_key": k,
            "slot_raw": v["slot_raw"],
            "courses": v["courses"],
            "divisions": v["divisions"],
            "merged_flag": v["merged_flag"],
            "students": students,
            "assigned": False
        })

    # Pre-compute total usable seats per session (same for all sessions unless room list changes)
    total_usable_all_rooms = int(sum(rooms_df["Seating Capacity"] // 2))

    # Prepare assignments list
    assignments = []  # list of dicts: {"day":d, "session":sess, "slots":[slot_objs]}

    # track division assignments to prevent clashes: division_taken[(day,session)][division]=True
    division_taken = defaultdict(lambda: defaultdict(bool))

    # iterate sessions chronologically
    slot_order = sorted(slots, key=lambda x: (x["slot_key"]))  # deterministic order; you can tweak priority here

    # For each day/session we will try to pack full slots (whole)
    day_idx = 1
    for day_i in range(num_days):
        for sess in ("FN", "AN"):
            if not sessions_per_day[day_i][sess]:
                continue
            remaining_capacity = total_usable_all_rooms
            placed_slots = []
            # iterate through unassigned slots and try to place them whole
            for slot in slot_order:
                if slot["assigned"]:
                    continue
                # skip if students greater than total capacity (cannot be placed in any session) — leave unassigned
                if slot["students"] > total_usable_all_rooms:
                    # cannot place in any session; skip (will remain unassigned)
                    continue
                # check division conflicts for this (day,session)
                conflict = False
                for div in slot["divisions"]:
                    if division_taken[(day_idx, sess)].get(div, False):
                        conflict = True
                        break
                if conflict:
                    continue
                # check if it fits in remaining capacity of this session
                if slot["students"] <= remaining_capacity:
                    # assign
                    placed_slots.append(slot)
                    remaining_capacity -= slot["students"]
                    slot["assigned"] = True
                    for div in slot["divisions"]:
                        division_taken[(day_idx, sess)][div] = True
                # if not fit, skip and try next slot (we don't partially place)
            assignments.append({"day": day_idx, "session": sess, "slots": placed_slots})
        day_idx += 1

    # Collect unassigned slots (students >0 and not assigned)
    unassigned = [s for s in slot_order if not s["assigned"]]

    return assignments, unassigned

# -------------------------
# Seating: create grids per room and seat assigned slots only (non-partial)
# (re-using earlier seating logic)
# -------------------------
def make_grid(rows, cols):
    return [["" for _ in range(cols)] for __ in range(rows)]

def allocate_seating_for_session(placed_slots, rooms_df, invigilators):
    """
    identical seating algorithm as before:
    - Build per-item list (merged -> single item; non-merged -> per-division items)
    - For each room: grid with rows=6 and cols=ceil(cap/6), usable seats = cap//2
    - Fill SLOT columns with blanks between; ensure consecutive SLOT columns have different base slotname
    - If only one remaining item and its base equals last placed base -> stop filling that room early and continue next room
    - Finally fallback-fill remaining seats ignoring base constraint
    """
    # Build parent groups and per-division items
    parent_groups = {}
    for slot in placed_slots:
        parent = slot["slot_key"]
        raw = slot.get("slot_raw", parent)
        base = raw.split("_Y")[0] if "_Y" in raw else (raw.split("_")[0] if "_" in raw else raw)
        if slot.get("merged_flag", False):
            items = [{"label_prefix": parent, "remaining": int(slot["students"])}]
        else:
            items = []
            for rr in sorted(slot["courses"], key=lambda x: x["DIVISION"]):
                div = rr["DIVISION"]
                items.append({"label_prefix": f"{parent}_{div}", "remaining": int(rr["NO_STUDENTS"])})
        parent_groups[parent] = {"base": base, "items": items}

    # active items flatten
    active_items = []
    for parent, v in parent_groups.items():
        for it in v["items"]:
            active_items.append({"parent": parent, "base": v["base"], "label_prefix": it["label_prefix"], "remaining": it["remaining"]})

    rooms = []
    for _, r in rooms_df.iterrows():
        cap = int(r["Seating Capacity"])
        rows = 6
        cols = max(1, math.ceil(cap / rows))
        rooms.append({
            "name": str(r["Room"]),
            "rows": rows,
            "cols": cols,
            "usable": cap // 2,
            "grid": make_grid(rows, cols),
            "invigilators": random.sample(invigilators, min(2, max(1, len(invigilators))))
        })

    placed_counters = defaultdict(int)

    # Fill rooms one by one strictly (no partial slot placement across a session)
    for room in rooms:
        if not any(it["remaining"] > 0 for it in active_items):
            break
        rows = room["rows"]
        cols = room["cols"]
        usable = room["usable"]
        placed_in_room = 0
        col_idx = 0
        last_slot_base = None

        # place while capacity and columns available
        while placed_in_room < usable and any(it["remaining"] > 0 for it in active_items) and col_idx < cols:
            # pick next item with remaining and base != last_slot_base
            chosen_idx = None
            for idx, it in enumerate(active_items):
                if it["remaining"] <= 0:
                    continue
                if last_slot_base is None or it["base"] != last_slot_base:
                    chosen_idx = idx
                    break
            if chosen_idx is None:
                # if only one item left and its base==last_slot_base -> stop this room (defer to next room)
                remaining_nonzero = [it for it in active_items if it["remaining"] > 0]
                if len(remaining_nonzero) <= 1:
                    break
                # else relax and pick any remaining
                for idx, it in enumerate(active_items):
                    if it["remaining"] > 0:
                        chosen_idx = idx
                        break
                if chosen_idx is None:
                    break

            chosen = active_items[chosen_idx]
            # place a full column (up to rows or remaining)
            for r_in in range(rows):
                if placed_in_room >= usable:
                    break
                if chosen["remaining"] <= 0:
                    room["grid"][r_in][col_idx] = ""
                    continue
                placed_counters[chosen["label_prefix"]] += 1
                labnum = placed_counters[chosen["label_prefix"]]
                room["grid"][r_in][col_idx] = f"{chosen['label_prefix']}-{labnum}"
                chosen["remaining"] -= 1
                placed_in_room += 1
            last_slot_base = chosen["base"]
            col_idx += 1
            # add blank column if space and not full
            if placed_in_room < usable and col_idx < cols:
                for r_in in range(rows):
                    room["grid"][r_in][col_idx] = ""
                col_idx += 1
        # end while for this room

    # Fallback fill remaining seats ignoring base constraint
    for room in rooms:
        rows = room["rows"]
        cols = room["cols"]
        usable = room["usable"]
        already = sum(1 for r in range(rows) for c in range(cols) if room["grid"][r][c])
        if already >= usable:
            continue
        for c in range(cols):
            for r_in in range(rows):
                if already >= usable:
                    break
                if room["grid"][r_in][c]:
                    continue
                found = False
                for it in active_items:
                    if it["remaining"] > 0:
                        placed_counters[it["label_prefix"]] += 1
                        labnum = placed_counters[it["label_prefix"]]
                        room["grid"][r_in][c] = f"{it['label_prefix']}-{labnum}"
                        it["remaining"] -= 1
                        already += 1
                        found = True
                        break
                if not found:
                    break
            if already >= usable:
                break

    return rooms

# -------------------------
# Write seating Excel
# -------------------------
def write_seating_excel(day_idx, placed_slots, rooms_df, invigilators):
    rooms_alloc = allocate_seating_for_session(placed_slots, rooms_df, invigilators)
    out_dir = Path("output/seating_arrangements")
    out_dir.mkdir(parents=True, exist_ok=True)
    outpath = out_dir / f"Day_{day_idx}.xlsx"
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for sess in ("FN", "AN"):
        ws = wb.create_sheet(sess)
        row_cursor = 1
        for room in rooms_alloc:
            ws.cell(row=row_cursor, column=1, value=f"Room: {room['name']}")
            ws.cell(row=row_cursor, column=3, value=f"Invigilators: {', '.join(room['invigilators'])}")
            row_cursor += 1
            for r in range(room["rows"]):
                for c in range(room["cols"]):
                    val = room["grid"][r][c] or ""
                    cell = ws.cell(row=row_cursor + r, column=c + 1, value=val)
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = border
            row_cursor += room["rows"] + 2
        # column widths
        max_cols = max((room["cols"] for room in rooms_alloc), default=1)
        for c in range(1, max_cols + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

    wb.save(outpath)
    print(f"Wrote seating for Day {day_idx}: {outpath}")

# -------------------------
# Timetable builder (from assignments)
# -------------------------
def build_timetable_from_assignments(df_courses, assignments, outpath):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    for sess in ("FN", "AN"):
        ws = wb.create_sheet(sess)
        ws.append(["Day", "Slots"])
        for alloc in assignments:
            if alloc["session"] != sess:
                continue
            names = [s["slot_key"] for s in alloc["slots"]]
            ws.append([alloc["day"], ", ".join(names)])
    ws_ref = wb.create_sheet("Reference Table")
    ws_ref.append(["ELECTIVE OR NOT", "FULLSEM OR HALFSEM", "SLOT NAME", "COURSE CODE", "COURSE TITLE", "DIVISION"])
    ref_df = df_courses[["ELECTIVE", "FULLSEM_TYPE", "SLOT", "COURSE_CODE", "COURSE_TITLE", "DIVISION"]].drop_duplicates()
    for _, r in ref_df.iterrows():
        ws_ref.append([r["ELECTIVE"], r["FULLSEM_TYPE"], r["SLOT"], r["COURSE_CODE"], r["COURSE_TITLE"], r["DIVISION"]])
    for sh in wb.sheetnames:
        ws = wb[sh]
        for col in ws.columns:
            first = col[0]
            try:
                ws.column_dimensions[first.column_letter].width = 25
            except Exception:
                pass
            for cell in col:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if cell.row == 1:
                    cell.font = Font(bold=True)
    Path(outpath).parent.mkdir(parents=True, exist_ok=True)
    wb.save(outpath)
    print(f"Wrote timetable: {outpath}")

# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    divisions, rooms_path, invig_path, num_days, sessions_per_day, num_years = get_user_inputs()
    df_courses = load_courses(divisions, num_years)
    if df_courses.empty:
        print("No courses found. Exiting.")
        raise SystemExit

    first_half, second_half = split_half(df_courses)

    rooms_df = pd.read_excel(rooms_path, engine="openpyxl")
    rooms_df.columns = [str(c).strip() for c in rooms_df.columns]
    if "Room" not in rooms_df.columns or "Seating Capacity" not in rooms_df.columns:
        raise ValueError("Rooms file must contain 'Room' and 'Seating Capacity' columns")

    inv_df = pd.read_excel(invig_path, engine="openpyxl", header=None)
    invig_list = [str(x).strip() for x in inv_df.iloc[:,0] if str(x).strip()]

    # Allocate slots by seating feasibility (first-half and second-half separately)
    assignments_first, unassigned_first = allocate_slots_by_seating_capacity(first_half, num_days, sessions_per_day, rooms_df)
    assignments_second, unassigned_second = allocate_slots_by_seating_capacity(second_half, num_days, sessions_per_day, rooms_df)

    # Combine assignments by day/session (we have a list in chronological order)
    # For building timetable files, we separate by half as you requested earlier
    build_timetable_from_assignments(first_half, assignments_first, "output/firsthalf_timetable.xlsx")
    build_timetable_from_assignments(second_half, assignments_second, "output/secondhalf_timetable.xlsx")

    # Write seating workbooks day-wise: for each day collect slots assigned in both halves and seat them
    # assignments_first/second contains entries with "day","session","slots"
    Path("output").mkdir(exist_ok=True)
    for day_idx in range(1, num_days + 1):
        # collect slots assigned to this day across both halves and sessions
        day_slots = []
        for alloc in assignments_first + assignments_second:
            if alloc["day"] == day_idx:
                day_slots.extend(alloc["slots"])
        if not day_slots:
            continue
        # seat them (function will create per-session sheets FN and AN but we pass full list)
        write_seating_excel(day_idx, day_slots, rooms_df, invig_list)

    # Output unassigned lists for your review
    if unassigned_first or unassigned_second:
        print("\nWARNING: Some slots could not be assigned in any session (too large or insufficient sessions).")
        print("Unassigned (first half):", [s["slot_key"] for s in unassigned_first])
        print("Unassigned (second half):", [s["slot_key"] for s in unassigned_second])
    else:
        print("\nAll slots assigned across given sessions/days.")

    print("\nDone. Files written to ./output and ./output/seating_arrangements")
