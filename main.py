import os
import math
import json
import random
import pandas as pd
from collections import defaultdict
from datetime import datetime, timedelta
from math import gcd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell


def load_settings(path="settings.json"):
    DEFAULT = {
        "working_days": ["Mon", "Tue", "Wed", "Thu", "Fri"],
        "working_hours": ["9:00", "18:30"],
        "break_slots": ["12:30-13:30", "16:30-17:00"],
        "slot_durations": {"lec": 1.5, "lab": 2.0, "tut": 1.0}
    }
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                data = json.load(f)
            print("✅ Loaded settings.json")
        except Exception as e:
            print("⚠️ Could not read settings.json, using defaults:", e)
            data = DEFAULT
    else:
        data = DEFAULT
        print("⚙️ Using default settings (settings.json not found)")

    breaks = []
    for b in data.get("break_slots", []):
        if isinstance(b, str) and "-" in b:
            a, c = b.split("-", 1)
            breaks.append((a.strip(), c.strip()))
        elif isinstance(b, (list, tuple)) and len(b) == 2:
            breaks.append((b[0], b[1]))
    data["break_slots"] = breaks

    print("Working days:", data["working_days"])
    print("Working hours:", data["working_hours"])
    print("Break slots:", data["break_slots"])
    print("Slot durations (hours):", data["slot_durations"])
    print("-" * 60)
    return data


def time_to_minutes(t):
    if isinstance(t, (int, float)):
        return int(t)
    h, m = map(int, str(t).split(":"))
    return h * 60 + m

def minutes_to_time(m):
    h = int(m // 60); mm = int(m % 60)
    return f"{h:02d}:{mm:02d}"

def gcd_list(nums):
    nums = [int(n) for n in nums if n and n > 0]
    if not nums:
        return 15
    g = nums[0]
    for n in nums[1:]:
        g = gcd(g, n)
    return g if g > 0 else 15


def parse_LTP_from_ltpsc(ltpsc):
    if pd.isna(ltpsc):
        return 0, 0, 0
    s = str(ltpsc).strip()
    parts = s.split("-")
    try:
        L = int(parts[0]) if len(parts) > 0 and parts[0] else 0
        T = int(parts[1]) if len(parts) > 1 and parts[1] else 0
        P = int(parts[2]) if len(parts) > 2 and parts[2] else 0
        return L, T, P
    except:
        return 0, 0, 0

def parse_people(s):
    if pd.isna(s) or str(s).strip() == "":
        return []
    return [x.strip() for x in str(s).split(",") if x.strip()]

def read_input_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def build_slot_requests_for_division(df, div_fullname, settings):
    normals = []
    baskets = {}

    for idx, row in df.iterrows():
        elective_flag = str(row.get("ELECTIVE OR NOT", "NO")).strip().upper()
        sem_type = str(row.get("FULLSEM OR HALFSEM", "fullsem")).strip().lower()
        code = str(row.get("COURSE CODE", "")).strip()
        title = str(row.get("COURSE TITLE", "")).strip()
        faculty = str(row.get("FACULTY", "")).strip()
        class_asst = parse_people(row.get("CLASS ASSISTANTS", ""))
        lab_asst = parse_people(row.get("LAB ASSISTANTS", ""))
        ltpsc = row.get("L-T-P-S-C", row.get("L-T-P", ""))
        L, T, P = parse_LTP_from_ltpsc(ltpsc)
        room_no = str(row.get("ROOM.NO", "")).strip()
        lab_room_no = str(row.get("LAB ROOM.NO", "")).strip()
        slot_base = str(row.get("SLOT NAME", "")).strip()
        merge_raw = str(row.get("MERGE", "")).strip()
        merge_with = [m.strip() for m in merge_raw.split(",") if m.strip()]

        for kind, hours in [("lec", L), ("tut", T), ("lab", P)]:
            if hours <= 0:
                continue
            dur_hours = settings["slot_durations"].get(kind, 1.0)

            full_slots = int(hours // dur_hours)
            remainder = hours % dur_hours

            for _ in range(full_slots):
                slot_label = f"{slot_base}-{kind}"
                occ = {
                    "slot_base": slot_base,
                    "slot_label": slot_label,
                    "code": code,
                    "title": title,
                    "faculty": faculty,
                    "class_asst": class_asst,
                    "lab_asst": lab_asst,
                    "L": L, "T": T, "P": P,
                    "L-T-P-S-C": ltpsc,
                    "room_no": room_no,
                    "lab_room_no": lab_room_no,
                    "sem_type": sem_type,
                    "merge_with": merge_with,
                    "division": div_fullname,
                    "kind": kind,
                    "_duration_hours": dur_hours
                }
                if elective_flag == "YES" and slot_base.lower().startswith("elective"):
                    baskets.setdefault(slot_base, []).append(occ)
                else:
                    occ["occ_id"] = f"{code}_{kind}_{random.randint(1000,9999)}"
                    normals.append(occ)

            if remainder > 0:
                slot_label = f"{slot_base}-{kind}"
                occ = {
                    "slot_base": slot_base,
                    "slot_label": slot_label,
                    "code": code,
                    "title": title,
                    "faculty": faculty,
                    "class_asst": class_asst,
                    "lab_asst": lab_asst,
                    "L": L, "T": T, "P": P,
                    "L-T-P-S-C": ltpsc,
                    "room_no": room_no,
                    "lab_room_no": lab_room_no,
                    "sem_type": sem_type,
                    "merge_with": merge_with,
                    "division": div_fullname,
                    "kind": kind,
                    "_duration_hours": remainder
                }
                if elective_flag == "YES" and slot_base.lower().startswith("elective"):
                    baskets.setdefault(slot_base, []).append(occ)
                else:
                    occ["occ_id"] = f"{code}_{kind}_{random.randint(1000,9999)}"
                    normals.append(occ)

    return normals, baskets



def schedule_globally(all_normals_per_div, all_baskets, settings, min_gap_minutes):
    days = settings["working_days"]
    wh_start = time_to_minutes(settings["working_hours"][0])
    wh_end = time_to_minutes(settings["working_hours"][1])

    dur_minutes = {k: int(v * 60) for k, v in settings["slot_durations"].items()}
    base_interval = gcd_list(list(dur_minutes.values()))
    if base_interval < 5:
        base_interval = 15

    interval_times = []
    t = wh_start
    while t < wh_end:
        interval_times.append(t)
        t += base_interval

    break_ranges = []
    for bstart, bend in settings.get("break_slots", []):
        bs = time_to_minutes(bstart); be = time_to_minutes(bend)
        break_ranges.append((bs, be))
    break_indices = set()
    for i, st in enumerate(interval_times):
        et = st + base_interval
        for bs, be in break_ranges:
            if not (et <= bs or st >= be):
                break_indices.add(i)

    occ_people = defaultdict(set)
    occ_rooms = defaultdict(set)
    placements = {div: {d: [] for d in days} for div in all_normals_per_div.keys()}
    unscheduled = []

    normal_list = []
    for div, slots in all_normals_per_div.items():
        for s in slots:
            entry = s.copy()
            entry["_division"] = div
            entry["_duration_min"] = dur_minutes.get(entry["kind"], 60)
            normal_list.append(entry)

    normal_list.sort(key=lambda x: (-x["_duration_min"], x.get("faculty", "")))
    fixed_lab_gap = 180

    def can_place_block(day, start_idx, n_intervals, busy_people, room, require_not_in_break=True):
        if start_idx + n_intervals > len(interval_times):
            return False
        for idx in range(start_idx, start_idx + n_intervals):
            if require_not_in_break and idx in break_indices:
                return False
            key = (day, idx)
            if busy_people & occ_people.get(key, set()):
                return False
            if room and room in occ_rooms.get(key, set()):
                return False
        return True

    def mark_block(day, start_idx, n_intervals, busy_people, room):
        for idx in range(start_idx, start_idx + n_intervals):
            key = (day, idx)
            occ_people[key].update(busy_people)
            if room:
                occ_rooms[key].add(room)

    for slot in normal_list:
        div = slot["_division"]
        duration = slot["_duration_min"]
        n_intervals = max(1, int(math.ceil(duration / base_interval)))

        busy_people = set()
        if slot.get("faculty"):
            busy_people.add(slot["faculty"])
        if slot["kind"] in ("lec", "tut"):
            busy_people.update(slot.get("class_asst", []))
            room = slot.get("room_no") or None
        else:
            busy_people.update(slot.get("lab_asst", []))
            room = slot.get("lab_room_no") or None

        placed = False
        days_to_try = days.copy()
        random.shuffle(days_to_try)
        for day in days_to_try:
            for start_idx in range(0, len(interval_times) - n_intervals + 1):
                if any(idx in break_indices for idx in range(start_idx, start_idx + n_intervals)):
                    continue

                violates_gap = False
                cand_start_min = interval_times[start_idx]
                cand_end_min = interval_times[start_idx + n_intervals - 1] + base_interval
                slot_base_curr = slot.get("slot_base", "")

                for (ex_start, ex_len, ex_label, ex_kind, ex_meta) in placements.get(div, {}).get(day, []):
                    ex_start_min = interval_times[ex_start]
                    ex_end_min = interval_times[ex_start + ex_len - 1] + base_interval
                    slot_base_existing = ex_meta.get("slot_base", "") if isinstance(ex_meta, dict) else ""

                    if not (cand_end_min + min_gap_minutes <= ex_start_min or cand_start_min >= ex_end_min + min_gap_minutes):
                        violates_gap = True
                        break

                    if slot_base_curr and slot_base_curr == slot_base_existing:
                        if abs(cand_start_min - ex_start_min) < 180:
                            violates_gap = True
                            break

                if violates_gap:
                    continue

                if slot["kind"] == "lab":
                    lab_check_intervals = int(math.ceil(fixed_lab_gap / base_interval))
                    low = max(0, start_idx - lab_check_intervals)
                    high = min(len(interval_times), start_idx + n_intervals + lab_check_intervals)
                    conflict_lab_gap = False
                    for idx in range(low, high):
                        key = (day, idx)
                        if busy_people & occ_people.get(key, set()):
                            conflict_lab_gap = True
                            break
                    if conflict_lab_gap:
                        continue

                if not can_place_block(day, start_idx, n_intervals, busy_people, room, require_not_in_break=True):
                    continue

                placements[div][day].append((start_idx, n_intervals, slot["slot_label"], slot["kind"], slot))
                mark_block(day, start_idx, n_intervals, busy_people, room)

                for mdiv in slot.get("merge_with", []):
                    mdiv = mdiv.strip()
                    if not mdiv or mdiv not in placements:
                        continue
                    placements[mdiv][day].append((start_idx, n_intervals, slot["slot_label"], slot["kind"], slot))
                    mark_block(day, start_idx, n_intervals, busy_people, room)

                placed = True
                break
            if placed:
                break

        if not placed:
            unscheduled.append(f"{slot.get('code') or slot.get('slot_label')} ({slot.get('kind')}) in {div} not placed")

    for basket_label, members in all_baskets.items():
        if not members:
            continue
        ref = members[0]
        kind = ref.get("kind", "lec")
        duration_min = dur_minutes.get(kind, 60)
        n_intervals = max(1, int(math.ceil(duration_min / base_interval)))

        combined_people = set()
        combined_rooms = set()
        basket_divs = set()
        for c in members:
            if c.get("faculty"):
                combined_people.add(c["faculty"])
            if c.get("kind") in ("lec", "tut"):
                combined_people.update(c.get("class_asst", []))
                if c.get("room_no"):
                    combined_rooms.add(c.get("room_no"))
            else:
                combined_people.update(c.get("lab_asst", []))
                if c.get("lab_room_no"):
                    combined_rooms.add(c.get("lab_room_no"))
            if c.get("division"):
                basket_divs.add(c.get("division"))
            for md in c.get("merge_with", []):
                if md and str(md).strip() and md in all_normals_per_div:
                    basket_divs.add(md.strip())

        placed = False
        days_to_try = days.copy()
        random.shuffle(days_to_try)
        for day in days_to_try:
            for start_idx in range(0, len(interval_times) - n_intervals + 1):
                if any(idx in break_indices for idx in range(start_idx, start_idx + n_intervals)):
                    continue
                conflict = False
                for idx in range(start_idx, start_idx + n_intervals):
                    key = (day, idx)
                    if combined_people & occ_people.get(key, set()):
                        conflict = True; break
                    if combined_rooms & occ_rooms.get(key, set()):
                        conflict = True; break
                if conflict:
                    continue
                for idx in range(start_idx, start_idx + n_intervals):
                    key = (day, idx)
                    occ_people[key].update(combined_people)
                    occ_rooms[key].update(combined_rooms)
                for mdiv in basket_divs:
                    if mdiv not in placements:
                        continue
                    placements[mdiv][day].append((start_idx, n_intervals, basket_label, kind, {"basket_members": members}))
                placed = True
                break
            if placed:
                break

        if not placed:
            unscheduled.append(f"Basket {basket_label} not placed")

    return placements, unscheduled, interval_times, base_interval, break_indices


def set_value_in_merged_region(ws, row, col_start, col_end, value):
    unmerge_ranges_overlapping(ws, row, col_start, col_end)

    if col_end > col_start:
        try:
            ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
        except Exception:
            pass

    try:
        ws.cell(row=row, column=col_start, value=value)
    except Exception:
        try:
            ws.unmerge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
        except Exception:
            pass
        try:
            ws.cell(row=row, column=col_start, value=value)
        except Exception:
            pass

def safe_sheet_title(raw, fallback_prefix="Sheet"):
    try:
        if pd.isna(raw): return None
    except Exception:
        pass
    s = str(raw) if raw is not None else ""
    s = s.strip()
    if s == "" or s.lower() == "nan": return None
    for ch in [":", "\\", "/", "?", "*", "[", "]"]:
        s = s.replace(ch, "_")
    if len(s) > 31: s = s[:31]
    return s

def write_year_excel(year, half_tag, placements, interval_times, base_interval, break_indices, colors, course_info_per_div, settings, outdir="timetable_outputs"):
    os.makedirs(outdir, exist_ok=True)
    fname = os.path.join(outdir, f"Timetable_Year{year}_{half_tag}.xlsx")
    wb = Workbook()
    try:
        wb.remove(wb.active)
    except Exception:
        pass
    days = settings["working_days"]

    time_headers = [f"{minutes_to_time(t)} - {minutes_to_time(t + base_interval)}" for t in interval_times]

    for div_index, (div, day_map) in enumerate(placements.items(), start=1):
        title_candidate = safe_sheet_title(div)
        if title_candidate is None:
            title_candidate = f"Div_{div_index}"
        base_title = title_candidate
        suffix = 1
        while title_candidate in wb.sheetnames:
            title_candidate = f"{base_title[:28]}_{suffix}"
            suffix += 1

        ws = wb.create_sheet(title=title_candidate)
        ws.append([f"Division: {div}    Year: {year}    Half: {half_tag}"])
        ws.append([])
        header = ["Day/Time"] + time_headers
        ws.append(header)
        header_row_idx = ws.max_row

        for day in days:
            ws.append([day] + [""] * len(time_headers))
        first_day_row = header_row_idx + 1

        for r_idx, day in enumerate(days):
            placements_for_day = day_map.get(day, [])
            placements_for_day.sort(key=lambda x: x[0])
            for (start_idx, n_intervals, slot_label, kind, meta) in placements_for_day:
                label = slot_label if slot_label else f"{meta.get('slot_base','')}-{kind}"
                excel_row = first_day_row + r_idx
                excel_col_start = 2 + start_idx
                excel_col_end = 2 + start_idx + n_intervals - 1

                set_value_in_merged_region(ws, excel_row, excel_col_start, excel_col_end, label)

                slot_base = meta.get("slot_base") if isinstance(meta, dict) else None
                if not slot_base:
                    if isinstance(slot_label, str) and "-" in slot_label:
                        slot_base = "-".join(slot_label.split("-")[:-1])
                    else:
                        slot_base = slot_label
                fill_color = colors.get(slot_base, "#DDDDDD")
                color_code = fill_color[1:] if str(fill_color).startswith("#") else str(fill_color)
                if not (isinstance(color_code, str) and len(color_code) in (6, 8) and all(ch in "0123456789ABCDEFabcdef" for ch in color_code)):
                    color_code = "DDDDDD"

                for ccol in range(excel_col_start, excel_col_end + 1):
                    c = ws.cell(row=excel_row, column=ccol)
                    c.fill = PatternFill(start_color=color_code, end_color=color_code, fill_type="solid")
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.font = Font(size=10, bold=True)

        for bi in sorted(break_indices):
            excel_col = 2 + bi
            for r in range(first_day_row, first_day_row + len(days)):
                try:
                    cell = ws.cell(row=r, column=excel_col)
                    cell.fill = PatternFill(start_color="EEEEEE", end_color="EEEEEE", fill_type="solid")
                    if not cell.value:
                        cell.value = "BREAK"
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                except Exception:
                    pass

        ws.column_dimensions[get_column_letter(1)].width = 14
        for ci in range(2, 2 + len(time_headers)):
            ws.column_dimensions[get_column_letter(ci)].width = 18

        ws.append([])
        ws.append(["Reference Table"])
        ws.append(["Slot Name", "Course Code", "Course Title", "Faculty", "L-T-P-S-C", "ROOM.NO", "LAB ROOM.NO"])
        seen = set()
        infos = course_info_per_div.get(div, [])
        slot_base_map = {}
        for info in infos:
            sb = info.get("slot_base") or ""
            if not sb:  # fallback
                sl = info.get("slot_label") or ""
                if "-" in sl:
                    sb = "-".join(sl.split("-")[:-1])
                else:
                    sb = sl
            if sb not in slot_base_map:
                slot_base_map[sb] = info

        for sb, info in slot_base_map.items():
            slot_label = sb
            code = info.get("code", "")
            ltp = info.get("L-T-P-S-C", "")
            ws.append([slot_label, code, info.get("title", ""), info.get("faculty", ""), ltp, info.get("room_no", ""), info.get("lab_room_no", "")])

    if not wb.sheetnames:
        wb.create_sheet(title="Timetable")
    wb.save(fname)
    print(f"✅ Saved: {fname}")

def ranges_overlap(a_start, a_end, b_start, b_end):
    return not (a_end < b_start or b_end < a_start)

def unmerge_ranges_overlapping(ws, row, col_start, col_end):
    to_unmerge = []
    for mr in list(ws.merged_cells.ranges):
        min_col, min_row, max_col, max_row = mr.min_col, mr.min_row, mr.max_col, mr.max_row
        if row >= min_row and row <= max_row:
            if ranges_overlap(col_start, col_end, min_col, max_col):
                to_unmerge.append(str(mr))
    for rng in to_unmerge:
        try:
            ws.unmerge_cells(rng)
        except:
            pass

def main():
    print("\n🧩 Timetable Generator\n")
    settings = load_settings()

    while True:
        try:
            min_gap = int(input("Enter minimum gap between consecutive slots in minutes (default 0): ") or "0")
            if min_gap < 0:
                print("Enter non-negative integer")
                continue
            break
        except Exception:
            print("Please enter integer minutes")

    print("Minimum gap:", min_gap, "minutes")
    print("-" * 60)

    n_years = int(input("Enter number of academic years: ").strip())
    inputs_per_year = {}
    for y in range(1, n_years + 1):
        n_div = int(input(f"Year {y}, number of divisions: ").strip())
        inputs_per_year[y] = {}
        for d in range(1, n_div + 1):
            full_div = input(f"  Full name for Division {d} (use exact name for MERGE, e.g. '1CSEA'): ").strip()
            if not full_div:
                full_div = input("    Division name cannot be blank — please enter full division name: ").strip()
            path = input(f"     Path to Excel/CSV for {full_div}: ").strip()
            inputs_per_year[y][full_div] = path

    for y in range(1, n_years + 1):
        print(f"\nProcessing Year {y} ...")
        div_paths = inputs_per_year[y]

        normals_first = {}
        normals_second = {}
        baskets_first = {}
        baskets_second = {}
        course_info_per_div = defaultdict(list)
        slot_bases_set = set()

        for div_full, path in div_paths.items():
            if not os.path.exists(path):
                print(f"⚠️ File not found: {path} for {div_full} — skipping division")
                continue
            df = read_input_file(path)
            normals, baskets = build_slot_requests_for_division(df, div_full, settings)

            normals_f = [n for n in normals if n.get("sem_type", "fullsem") in ("fullsem", "halfsem-1")]
            normals_s = [n for n in normals if n.get("sem_type", "fullsem") in ("fullsem", "halfsem-2") or n.get("sem_type", "fullsem") == "fullsem"]

            normals_first[div_full] = normals_f
            normals_second[div_full] = normals_s

            for b_label, members in baskets.items():
                sems = [m.get("sem_type", "fullsem") for m in members]
                if any(s in ("fullsem", "halfsem-1") for s in sems):
                    baskets_first.setdefault(b_label, []).extend(members)
                if any(s in ("fullsem", "halfsem-2") for s in sems):
                    baskets_second.setdefault(b_label, []).extend(members)

            for n in normals:
                course_info_per_div[div_full].append(n)
                sb = n.get("slot_base") or ""
                if sb:
                    slot_bases_set.add(sb)
                else:
                    sl = n.get("slot_label") or ""
                    if "-" in sl:
                        slot_bases_set.add("-".join(sl.split("-")[:-1]))
                    else:
                        slot_bases_set.add(sl)

            for blist in baskets.values():
                for c in blist:
                    course_info_per_div[div_full].append(c)
                    sb = c.get("slot_base") or ""
                    if sb:
                        slot_bases_set.add(sb)

        colors = {}
        palette = [
            "FF5733", "FF8D1A", "FFC300", "FFEA00", "9AFB60", "2ECC71", "27AE60",
            "00B2FF", "3498DB", "6C5CE7", "9B59B6", "F06292", "FFB6C1", "FF7F50",
            "D35400", "E67E22", "F39C12", "F1C40F", "2ECC71", "1ABC9C"
        ]
        random.seed(42)
        slot_bases_sorted = sorted(list(slot_bases_set))
        for s in slot_bases_sorted:
            colors[s] = "#" + random.choice(palette)

        # schedule first half
        placements_first, uns_first, interval_times, base_interval, break_indices = schedule_globally(normals_first, baskets_first, settings, min_gap)
        write_year_excel(y, "first_halfsem", placements_first, interval_times, base_interval, break_indices, colors, course_info_per_div, settings)

        # schedule second half
        placements_second, uns_second, interval_times2, base_interval2, break_indices2 = schedule_globally(normals_second, baskets_second, settings, min_gap)
        write_year_excel(y, "second_halfsem", placements_second, interval_times2, base_interval2, break_indices2, colors, course_info_per_div, settings)

        uns_total = uns_first + uns_second
        if uns_total:
            print("\n⚠️ Unscheduled (some duplicates possible):")
            for u in uns_total[:200]:
                print("   ", u)
            if len(uns_total) > 200:
                print("   ...", len(uns_total) - 200, "more not shown ...")

    print("\n✅ All done. Timetables saved in ./timetable_outputs")

if __name__ == "__main__":
    main()
