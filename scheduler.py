import pandas as pd
from datetime import datetime, timedelta
import math

SHIFTS_PER_DAY   = 3
BATCHES_PER_SHIFT = 2


def shift_index(date_str, shift):
    """Convert (date, shift) to a single integer index."""
    d     = datetime.strptime(str(date_str), "%Y-%m-%d")
    epoch = datetime(2000, 1, 1)
    return (d - epoch).days * SHIFTS_PER_DAY + (int(shift) - 1)


def index_to_shift(idx):
    """Convert integer index back to (date_str, shift)."""
    epoch   = datetime(2000, 1, 1)
    d       = epoch + timedelta(days=idx // SHIFTS_PER_DAY)
    shift   = (idx % SHIFTS_PER_DAY) + 1
    return d.strftime("%Y-%m-%d"), shift


def same_week(date_str1, date_str2):
    d1 = datetime.strptime(str(date_str1), "%Y-%m-%d")
    d2 = datetime.strptime(str(date_str2), "%Y-%m-%d")
    return d1.isocalendar()[1] == d2.isocalendar()[1] and d1.year == d2.year


def generate_mixing_schedule(master_mixer, master_produk, filling_plan):
    """
    Generate mixing schedule from master data and filling plan.

    Rules:
    - Mixing must finish at least 1 shift before filling
    - 1 shift = 2 batches per mixer
    - Changing product group on same mixer = 1 shift cleaning
    - Urgent: filling date locked
    - Not urgent: filling date can shift within same ISO week
    - Planner inputs in CS; converted to kg via Kg_per_CS
    """

    warnings      = []
    shifted       = []
    unscheduled   = []
    schedule_rows = []

    # ── Prep master data ──────────────────────────────────────
    mixer_df  = master_mixer.copy()
    produk_df = master_produk.copy()
    plan_df   = filling_plan.copy()

    produk_df["Mixer_List"] = produk_df["Mixer_Kompatibel"].apply(
        lambda x: [m.strip() for m in str(x).split(",")]
    )

    # ── Mixer state tracker ───────────────────────────────────
    mixer_schedule  = {row["Mixer"]: {} for _, row in mixer_df.iterrows()}
    mixer_last_grup = {row["Mixer"]: None for _, row in mixer_df.iterrows()}

    def get_mixer_capacity(mixer_name):
        row = mixer_df[mixer_df["Mixer"] == mixer_name]
        return float(row["Kapasitas_kg"].values[0]) if not row.empty else 0

    def get_shift_state(mixer_name, sidx):
        if sidx not in mixer_schedule[mixer_name]:
            mixer_schedule[mixer_name][sidx] = {
                "batches_used": 0, "grup": None,
                "items": [], "cleaning": False
            }
        return mixer_schedule[mixer_name][sidx]

    def mark_cleaning(mixer_name, sidx):
        state = get_shift_state(mixer_name, sidx)
        state["cleaning"]      = True
        state["batches_used"]  = BATCHES_PER_SHIFT

    def book_batches(mixer_name, sidx, n_batches, grup, kode, nama, kg_per_batch, total_kg, total_cs):
        state = get_shift_state(mixer_name, sidx)
        state["batches_used"] += n_batches
        state["grup"]          = grup
        state["items"].append({
            "kode": kode, "nama": nama,
            "batches": n_batches,
            "kg_per_batch": kg_per_batch,
            "total_kg": total_kg,
            "total_cs": total_cs
        })
        mixer_last_grup[mixer_name] = grup

    # ── Sort plan: urgent first, then by filling shift ────────
    plan_df["_sidx"]         = plan_df.apply(lambda r: shift_index(r["Tanggal_Filling"], r["Shift_Filling"]), axis=1)
    plan_df["_urgent_sort"]  = plan_df["Urgent"].apply(lambda x: 0 if x == "Urgent" else 1)
    plan_df = plan_df.sort_values(["_urgent_sort", "_sidx"]).reset_index(drop=True)

    # ── Schedule each item ────────────────────────────────────
    for _, item in plan_df.iterrows():
        kode        = item["Kode_Produk"]
        nama        = item.get("Nama_Produk", kode)
        target_cs   = float(item["Target_CS"])
        fill_date   = str(item["Tanggal_Filling"])
        fill_shift  = int(item["Shift_Filling"])
        is_urgent   = item["Urgent"] == "Urgent"

        # Lookup produk
        prod_row = produk_df[produk_df["Kode_Produk"] == kode]
        if prod_row.empty:
            unscheduled.append(f"Produk {kode} tidak ditemukan di Master Produk.")
            continue

        kg_per_cs    = float(prod_row["Kg_per_CS"].values[0])
        target_kg    = target_cs * kg_per_cs
        grup_produk  = prod_row["Grup_Cleaning"].values[0]
        mixer_compat = prod_row["Mixer_List"].values[0]

        # Candidate filling slots (not urgent = can shift within week)
        candidate_slots = [(fill_date, fill_shift)]
        if not is_urgent:
            d = datetime.strptime(fill_date, "%Y-%m-%d")
            for delta in range(0, 7):
                nd = d + timedelta(days=delta)
                nd_str = nd.strftime("%Y-%m-%d")
                if not same_week(fill_date, nd_str):
                    break
                start_s = fill_shift + 1 if delta == 0 else 1
                for s in range(start_s, SHIFTS_PER_DAY + 1):
                    candidate_slots.append((nd_str, s))

        scheduled = False

        for try_fill_date, try_fill_shift in candidate_slots:
            try_deadline = shift_index(try_fill_date, try_fill_shift) - 1
            temp_remaining  = target_kg
            temp_assignments = []
            temp_search      = try_deadline

            while temp_remaining > 0:
                if temp_search < 0:
                    break

                s_date, s_shift = index_to_shift(temp_search)

                for mixer_name in mixer_compat:
                    if mixer_name not in mixer_schedule:
                        continue

                    cap   = get_mixer_capacity(mixer_name)
                    state = mixer_schedule[mixer_name].get(temp_search, {
                        "batches_used": 0, "grup": None,
                        "cleaning": False, "items": []
                    })

                    if state.get("cleaning", False):
                        continue

                    avail = BATCHES_PER_SHIFT - state.get("batches_used", 0)
                    if avail <= 0:
                        continue

                    # Check if cleaning needed before this shift
                    used_shifts  = sorted([s for s in mixer_schedule[mixer_name] if s < temp_search], reverse=True)
                    last_grup    = mixer_schedule[mixer_name][used_shifts[0]]["grup"] if used_shifts else None
                    needs_clean  = last_grup is not None and last_grup != grup_produk

                    if needs_clean:
                        continue

                    use_kg      = min(avail * cap, temp_remaining)
                    use_batches = math.ceil(use_kg / cap)
                    actual_kg   = use_batches * cap
                    actual_cs   = actual_kg / kg_per_cs

                    temp_assignments.append({
                        "mixer":       mixer_name,
                        "shift_idx":   temp_search,
                        "batches":     use_batches,
                        "kg_per_batch": cap,
                        "kg":          actual_kg,
                        "cs":          actual_cs,
                        "needs_clean": needs_clean,
                        "last_grup":   last_grup,
                        "date":        s_date,
                        "shift":       s_shift
                    })
                    temp_remaining -= use_kg
                    break

                temp_search -= 1

            if temp_remaining > 0:
                continue

            # ── Commit assignments ────────────────────────────
            for a in temp_assignments:
                mixer_name = a["mixer"]
                sidx       = a["shift_idx"]

                # Insert cleaning shift if needed
                used_shifts = sorted([s for s in mixer_schedule[mixer_name] if s < sidx], reverse=True)
                last_grup   = mixer_schedule[mixer_name][used_shifts[0]]["grup"] if used_shifts else None
                if last_grup is not None and last_grup != grup_produk:
                    clean_idx = sidx - 1
                    mark_cleaning(mixer_name, clean_idx)
                    cd, cs_shift = index_to_shift(clean_idx)
                    schedule_rows.append({
                        "Tanggal":         cd,
                        "Shift":           cs_shift,
                        "Mixer":           mixer_name,
                        "Produk":          "— CLEANING —",
                        "Kode_Produk":     "",
                        "Batches":         "-",
                        "Kapasitas_Mixer": get_mixer_capacity(mixer_name),
                        "Total_CS":        0,
                        "Total_kg":        0,
                        "Cleaning":        True
                    })

                book_batches(mixer_name, sidx, a["batches"], grup_produk,
                             kode, nama, a["kg_per_batch"], a["kg"], a["cs"])

                schedule_rows.append({
                    "Tanggal":         a["date"],
                    "Shift":           a["shift"],
                    "Mixer":           mixer_name,
                    "Produk":          nama,
                    "Kode_Produk":     kode,
                    "Batches":         a["batches"],
                    "Kapasitas_Mixer": a["kg_per_batch"],
                    "Total_CS":        round(a["cs"], 2),
                    "Total_kg":        round(a["kg"], 2),
                    "Cleaning":        False
                })

            # Track if filling was shifted
            if try_fill_date != fill_date or try_fill_shift != fill_shift:
                shifted.append({
                    "Kode_Produk":  kode,
                    "Nama_Produk":  nama,
                    "Target_CS":    target_cs,
                    "Filling_Asal": f"{fill_date} Shift {fill_shift}",
                    "Filling_Baru": f"{try_fill_date} Shift {try_fill_shift}",
                    "Alasan":       "Kapasitas mixer penuh"
                })

            scheduled = True
            break

        if not scheduled:
            unscheduled.append(
                f"{kode} - {nama}: Tidak bisa dijadwalkan "
                f"(target {target_cs} CS / {target_kg} kg, "
                f"filling {fill_date} Shift {fill_shift})"
            )

    # ── Build output dataframe ────────────────────────────────
    if schedule_rows:
        schedule_df = pd.DataFrame(schedule_rows)
        schedule_df = schedule_df.sort_values(["Tanggal", "Shift", "Mixer"]).reset_index(drop=True)
        schedule_df = schedule_df[[
            "Tanggal", "Shift", "Mixer", "Produk", "Kode_Produk",
            "Batches", "Kapasitas_Mixer", "Total_CS", "Total_kg", "Cleaning"
        ]]
    else:
        schedule_df = pd.DataFrame()

    return {
        "schedule":   schedule_df,
        "warnings":   warnings,
        "shifted":    shifted,
        "unscheduled": unscheduled
    }
