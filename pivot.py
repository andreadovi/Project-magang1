import pandas as pd
import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DAYS_ID = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

def date_to_day(date_str):
    return DAYS_ID[datetime.strptime(str(date_str), "%Y-%m-%d").weekday()]


def build_pivot(schedule_df, master_mixer, master_produk, date_range):
    """
    Build pivot: rows = (Mixer, Kode_Produk, Nama_Produk)
                 cols = per date x 3 shifts
                 values = Total_kg
                 cleaning slots = marked separately
    """
    if schedule_df.empty:
        return pd.DataFrame(), {}

    col_keys   = []  # (date_str, shift)
    col_labels = []  # "Kamis, 07 Jun\nShift 1"
    for d in date_range:
        day_name = date_to_day(d)
        dt       = datetime.strptime(d, "%Y-%m-%d")
        date_lbl = dt.strftime("%d %b")
        for s in [1, 2, 3]:
            col_keys.append((d, s))
            col_labels.append(f"{day_name}, {date_lbl}\nShift {s}")

    # ── Row index: per mixer → per compatible produk ──────────
    mixer_order = list(master_mixer["Mixer"])
    rows = []
    for mixer in mixer_order:
        compat = master_produk[
            master_produk["Mixer_Kompatibel"].str.contains(mixer, na=False)
        ][["Kode_Produk", "Nama_Produk"]].drop_duplicates()
        for _, p in compat.iterrows():
            rows.append((mixer, p["Kode_Produk"], p["Nama_Produk"]))

    # ── Fill pivot data ───────────────────────────────────────
    pivot_data     = {}  # (mixer, kode) -> {(date, shift): kg}
    cleaning_cells = set()  # (mixer, kode, date, shift)
    resting_cells  = set()  # (mixer, kode, date, shift) — for info only
    scheduled_mixer = {}  # kode -> actual mixer used (from schedule_df directly)

    for _, row in schedule_df.iterrows():
        mx       = row["Mixer"]
        date     = row["Tanggal"]
        shift    = int(row["Shift"])
        cleaning = row.get("Cleaning", False)

        if cleaning:
            for r in rows:
                if r[0] == mx:
                    cleaning_cells.add((mx, r[1], date, shift))
            continue

        kode         = row["Kode_Produk"]
        kg           = float(row["Total_kg"]) if row["Total_kg"] else 0
        resting_days = int(row.get("Resting_Days", 0))

        # Track which mixer actually used for this product
        scheduled_mixer[kode] = mx

        key = (mx, kode)
        if key not in pivot_data:
            pivot_data[key] = {}
        pivot_data[key][(date, shift)] = pivot_data[key].get((date, shift), 0) + kg

        # Mark resting period cells (for visual info)
        if resting_days > 0:
            mix_dt = datetime.strptime(date, "%Y-%m-%d")
            for rd in range(1, resting_days + 1):
                rest_dt  = mix_dt + timedelta(days=rd)
                rest_str = rest_dt.strftime("%Y-%m-%d")
                for rs in [1, 2, 3]:
                    resting_cells.add((mx, kode, rest_str, rs))

    # ── Build dataframe ───────────────────────────────────────
    records = []
    for row_mixer, kode, nama in rows:
        # Skip if product was scheduled on a different mixer
        if kode in scheduled_mixer and scheduled_mixer[kode] != row_mixer:
            continue
        # Skip if product has no data AND no cleaning on this mixer
        has_data  = bool(pivot_data.get((row_mixer, kode), {}))
        has_clean = any((row_mixer, kode, d, s) in cleaning_cells for (d, s) in col_keys)
        if not has_data and not has_clean:
            continue

        rec = {"Mixer": row_mixer, "Kode_Produk": kode, "Nama_Produk": nama}
        for (d, s), label in zip(col_keys, col_labels):
            val = pivot_data.get((row_mixer, kode), {}).get((d, s), "")
            rec[label] = val if val != 0 else ""
        records.append(rec)

    pivot_df = pd.DataFrame(records)

    return pivot_df, {
        "col_keys":       col_keys,
        "col_labels":     col_labels,
        "cleaning_cells": cleaning_cells,
        "resting_cells":  resting_cells,
        "rows":           rows
    }


def pivot_to_excel(pivot_df, meta, master_mixer):
    if pivot_df.empty:
        buf = io.BytesIO()
        Workbook().save(buf)
        return buf.getvalue()

    wb = Workbook()
    ws = wb.active
    ws.title = "Jadwal Mixing"

    col_keys       = meta["col_keys"]
    cleaning       = meta["cleaning_cells"]
    resting        = meta["resting_cells"]
    rows           = meta["rows"]
    DATA_START_COL = 4

    # ── Fills & fonts ─────────────────────────────────────────
    hdr_fill    = PatternFill("solid", fgColor="1F4E79")
    hdr_font    = Font(bold=True, color="FFFFFF")
    subhdr_fill = PatternFill("solid", fgColor="2E75B6")
    subhdr_font = Font(bold=True, color="FFFFFF")
    clean_fill  = PatternFill("solid", fgColor="BDD7EE")   # biru = cleaning
    rest_fill   = PatternFill()  # no resting color
    mixer_fill  = PatternFill("solid", fgColor="D9E1F2")
    alt_fill    = PatternFill("solid", fgColor="EBF1DE")
    empty_fill  = PatternFill("solid", fgColor="FFFFFF")
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin        = Side(style="thin", color="BFBFBF")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Row 1: fixed headers (merged 2 rows) ──────────────────
    for col, label in [(1, "Mixer"), (2, "Kode\nProduk"), (3, "Nama Produk")]:
        cell            = ws.cell(row=1, column=col, value=label)
        cell.fill       = hdr_fill
        cell.font       = hdr_font
        cell.alignment  = center
        ws.merge_cells(start_row=1, end_row=2, start_column=col, end_column=col)

    # ── Row 1: day headers (merged per 3 shifts) ──────────────
    date_groups = {}
    for i, (d, s) in enumerate(col_keys):
        date_groups.setdefault(d, []).append(DATA_START_COL + i)

    # ── Single row header: "Kamis, 07 Jun | Shift 1" ──────────
    # (no merge needed — each date is unique)
    for i, (d, s) in enumerate(col_keys):
        col      = DATA_START_COL + i
        day_name = date_to_day(d)
        dt_lbl   = datetime.strptime(d, "%Y-%m-%d").strftime("%d %b")
        cell     = ws.cell(row=1, column=col, value=f"{day_name}, {dt_lbl}")
        cell.fill      = subhdr_fill
        cell.font      = subhdr_font
        cell.alignment = center
        ws.merge_cells(start_row=1, end_row=1, start_column=col, end_column=col)

    # ── Row 2: shift headers ──────────────────────────────────
    for i, (d, s) in enumerate(col_keys):
        col        = DATA_START_COL + i
        cell       = ws.cell(row=2, column=col, value=f"Shift {s}")
        cell.fill  = subhdr_fill
        cell.font  = subhdr_font
        cell.alignment = center

    # ── Data rows ─────────────────────────────────────────────
    mixer_list = list(master_mixer["Mixer"])
    cur_row    = 3

    for mixer in mixer_list:
        mixer_rows = [(m, k, n) for m, k, n in rows if m == mixer]
        if not mixer_rows:
            continue
        mixer_start = cur_row

        for ri, (m, kode, nama) in enumerate(mixer_rows):
            ws.cell(row=cur_row, column=1, value=mixer if ri == 0 else "")
            ws.cell(row=cur_row, column=2, value=kode).alignment = center
            ws.cell(row=cur_row, column=3, value=nama).alignment = center

            row_bg = alt_fill if ri % 2 == 0 else empty_fill

            for i, (d, s) in enumerate(col_keys):
                col  = DATA_START_COL + i
                cell = ws.cell(row=cur_row, column=col)
                cell.alignment = center
                cell.border    = border

                if (mixer, kode, d, s) in cleaning:
                    cell.fill  = clean_fill
                    cell.value = ""
                elif False:  # resting color disabled
                    cell.fill  = rest_fill
                    cell.value = ""
                else:
                    dt_lbl = datetime.strptime(d, "%Y-%m-%d").strftime("%d %b")
                    label  = f"{date_to_day(d)}, {dt_lbl}\nShift {s}"
                    if label in pivot_df.columns:
                        val_series = pivot_df.loc[
                            (pivot_df["Mixer"] == row_mixer) &
                            (pivot_df["Kode_Produk"] == kode), label
                        ]
                        v = val_series.values[0] if len(val_series) > 0 else ""
                    else:
                        v = ""
                    cell.value = v if v != "" else ""
                    cell.fill  = row_bg

            cur_row += 1

        # Merge mixer column
        if cur_row - 1 > mixer_start:
            ws.merge_cells(start_row=mixer_start, end_row=cur_row - 1,
                           start_column=1, end_column=1)
        cell_mixer           = ws.cell(row=mixer_start, column=1)
        cell_mixer.fill      = mixer_fill
        cell_mixer.font      = Font(bold=True)
        cell_mixer.alignment = center

    # ── Legend row ────────────────────────────────────────────
    cur_row += 1
    ws.cell(row=cur_row, column=1, value="Keterangan:").font = Font(bold=True)
    legends = [
        (clean_fill, "Cleaning (ganti grup produk)"),
        (rest_fill,  "Resting period (produk perlu didiamkan 2 hari)")
    ]
    for i, (fill, label) in enumerate(legends):
        col  = 2 + i * 2
        cell = ws.cell(row=cur_row, column=col)
        cell.fill = fill
        ws.cell(row=cur_row, column=col + 1, value=label)

    # ── Column widths ─────────────────────────────────────────
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 22
    for i in range(len(col_keys)):
        ws.column_dimensions[get_column_letter(DATA_START_COL + i)].width = 10

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
