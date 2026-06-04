import pandas as pd
from datetime import datetime, timedelta
import math

SHIFTS_PER_DAY    = 3
BATCHES_PER_SHIFT = 2


def shift_index(date_str, shift):
    d     = datetime.strptime(str(date_str), "%Y-%m-%d")
    epoch = datetime(2000, 1, 1)
    return (d - epoch).days * SHIFTS_PER_DAY + (int(shift) - 1)


def index_to_shift(idx):
    epoch = datetime(2000, 1, 1)
    d     = epoch + timedelta(days=idx // SHIFTS_PER_DAY)
    shift = (idx % SHIFTS_PER_DAY) + 1
    return d.strftime("%Y-%m-%d"), shift


def same_week(date_str1, date_str2):
    d1 = datetime.strptime(str(date_str1), "%Y-%m-%d")
    d2 = datetime.strptime(str(date_str2), "%Y-%m-%d")
    return d1.isocalendar()[1] == d2.isocalendar()[1] and d1.year == d2.year


def get_deadline_idx(fill_date, fill_shift, resting_days):
    """Latest allowed mixing shift index before filling."""
    if resting_days == 0:
        return shift_index(fill_date, fill_shift) - 1
    else:
        fill_dt  = datetime.strptime(str(fill_date), "%Y-%m-%d")
        mix_date = fill_dt - timedelta(days=resting_days)
        return shift_index(mix_date.strftime("%Y-%m-%d"), SHIFTS_PER_DAY)


def try_schedule_on_mixer(mixer_name, mixer_schedule, deadline_idx,
                           target_kg, cap, grup_produk, mixer_last_grup):
    """
    Try to schedule all target_kg on a SINGLE mixer using CONTIGUOUS slots
    going backwards from deadline_idx. No gaps allowed.
    Slots with partial availability are included; fully blocked slots
    (cleaning or grup conflict) cause the window to shift back.
    """
    remaining_kg  = target_kg
    search_idx    = deadline_idx
    assignments   = []

    while remaining_kg > 0:
        if search_idx < 0:
            return None

        state = mixer_schedule[mixer_name].get(search_idx, {
            "batches_used": 0, "grup": None,
            "cleaning": False, "items": []
        })

        # Hard block: cleaning shift
        if state.get("cleaning", False):
            # Gap not allowed — reset and move window before this block
            assignments  = []
            remaining_kg = target_kg
            search_idx  -= 1
            continue

        avail = BATCHES_PER_SHIFT - state.get("batches_used", 0)

        # Hard block: fully used
        if avail <= 0:
            assignments  = []
            remaining_kg = target_kg
            search_idx  -= 1
            continue

        # Hard block: grup conflict
        used_before = sorted([s for s in mixer_schedule[mixer_name] if s < search_idx], reverse=True)
        last_grup   = mixer_schedule[mixer_name][used_before[0]]["grup"] if used_before else None
        if last_grup is not None and last_grup != grup_produk:
            assignments  = []
            remaining_kg = target_kg
            search_idx  -= 1
            continue

        # Slot is usable — add to front of assignments (we go backwards)
        use_kg      = min(avail * cap, remaining_kg)
        use_batches = math.ceil(use_kg / cap)
        actual_kg   = use_batches * cap
        s_date, s_shift = index_to_shift(search_idx)
        assignments.insert(0, {
            "mixer":        mixer_name,
            "shift_idx":    search_idx,
            "batches":      use_batches,
            "kg_per_batch": cap,
            "kg":           actual_kg,
            "cs":           0,
            "last_grup":    last_grup,
            "date":         s_date,
            "shift":        s_shift
        })
        remaining_kg -= actual_kg
        search_idx   -= 1

    return assignments if remaining_kg <= 0 else None


def generate_mixing_schedule(master_mixer, master_produk, filling_plan):
    warnings      = []
    shifted       = []
    unscheduled   = []
    schedule_rows = []

    mixer_df  = master_mixer.copy()
    produk_df = master_produk.copy()
    plan_df   = filling_plan.copy()

    produk_df["Mixer_List"] = produk_df["Mixer_Kompatibel"].apply(
        lambda x: [m.strip() for m in str(x).split(",")]
    )
    if "Resting_Days" not in produk_df.columns:
        produk_df["Resting_Days"] = 0
    produk_df["Resting_Days"] = produk_df["Resting_Days"].fillna(0).astype(int)

    # ── Mixer state ───────────────────────────────────────────
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
        state["cleaning"]     = True
        state["batches_used"] = BATCHES_PER_SHIFT

    def book_batches(mixer_name, sidx, n_batches, grup, kode, nama,
                     kg_per_batch, total_kg, total_cs):
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

    # ── Sort: urgent first, then by filling shift ─────────────
    plan_df["_sidx"]        = plan_df.apply(
        lambda r: shift_index(r["Tanggal_Filling"], r["Shift_Filling"]), axis=1)
    plan_df["_urgent_sort"] = plan_df["Urgent"].apply(
        lambda x: 0 if x == "Urgent" else 1)
    plan_df = plan_df.sort_values(["_urgent_sort", "_sidx"]).reset_index(drop=True)

    # ── Schedule each item ────────────────────────────────────
    for _, item in plan_df.iterrows():
        kode       = item["Kode_Produk"]
        nama       = item.get("Nama_Produk", kode)
        target_cs  = float(item["Target_CS"])
        fill_date  = str(item["Tanggal_Filling"])
        fill_shift = int(item["Shift_Filling"])
        is_urgent  = item["Urgent"] == "Urgent"

        prod_row = produk_df[produk_df["Kode_Produk"].astype(str).str.strip() == str(kode).strip()]
        if prod_row.empty:
            unscheduled.append(f"Produk {kode} tidak ditemukan di Master Produk.")
            continue

        kg_per_cs    = float(prod_row["Kg_per_CS"].values[0])
        target_kg    = target_cs * kg_per_cs
        grup_produk  = prod_row["Grup_Cleaning"].values[0]
        mixer_compat = prod_row["Mixer_List"].values[0]
        resting_days = int(prod_row["Resting_Days"].values[0])

        # Candidate filling slots
        candidate_slots = [(fill_date, fill_shift)]
        if not is_urgent:
            d = datetime.strptime(fill_date, "%Y-%m-%d")
            for delta in range(0, 7):
                nd     = d + timedelta(days=delta)
                nd_str = nd.strftime("%Y-%m-%d")
                if not same_week(fill_date, nd_str):
                    break
                start_s = fill_shift + 1 if delta == 0 else 1
                for s in range(start_s, SHIFTS_PER_DAY + 1):
                    candidate_slots.append((nd_str, s))

        scheduled = False

        for try_fill_date, try_fill_shift in candidate_slots:
            try_deadline = get_deadline_idx(try_fill_date, try_fill_shift, resting_days)

            # ── Try each compatible mixer independently ────────
            best_assignments = None
            chosen_mixer     = None

            for mixer_name in mixer_compat:
                if mixer_name not in mixer_schedule:
                    continue

                cap  = get_mixer_capacity(mixer_name)
                if cap <= 0:
                    continue

                assignments = try_schedule_on_mixer(
                    mixer_name, mixer_schedule, try_deadline,
                    target_kg, cap, grup_produk, mixer_last_grup
                )

                if assignments is not None:
                    # Prefer mixer with fewest total shifts used (least busy)
                    if best_assignments is None or \
                       len(assignments) < len(best_assignments):
                        best_assignments = assignments
                        chosen_mixer     = mixer_name

            if best_assignments is None:
                continue  # try next filling slot

            # ── Commit assignments ────────────────────────────
            for a in best_assignments:
                mixer_name = chosen_mixer
                sidx       = a["shift_idx"]

                # Insert cleaning if needed
                used_before = sorted([s for s in mixer_schedule[mixer_name] if s < sidx], reverse=True)
                last_grup   = mixer_schedule[mixer_name][used_before[0]]["grup"] if used_before else None
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
                        "Cleaning":        True,
                        "Resting_Days":    0
                    })

                actual_cs = round(a["kg"] / kg_per_cs, 2)
                book_batches(mixer_name, sidx, a["batches"], grup_produk,
                             kode, nama, a["kg_per_batch"], a["kg"], actual_cs)

                s_date, s_shift = a["date"], a["shift"]
                schedule_rows.append({
                    "Tanggal":         s_date,
                    "Shift":           s_shift,
                    "Mixer":           mixer_name,
                    "Produk":          nama,
                    "Kode_Produk":     kode,
                    "Batches":         a["batches"],
                    "Kapasitas_Mixer": a["kg_per_batch"],
                    "Total_CS":        actual_cs,
                    "Total_kg":        round(a["kg"], 2),
                    "Cleaning":        False,
                    "Resting_Days":    resting_days
                })

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

    if schedule_rows:
        schedule_df = pd.DataFrame(schedule_rows)
        schedule_df = schedule_df.sort_values(["Tanggal", "Shift", "Mixer"]).reset_index(drop=True)
        schedule_df = schedule_df[[
            "Tanggal", "Shift", "Mixer", "Produk", "Kode_Produk",
            "Batches", "Kapasitas_Mixer", "Total_CS", "Total_kg",
            "Resting_Days", "Cleaning"
        ]]
    else:
        schedule_df = pd.DataFrame()

    return {
        "schedule":    schedule_df,
        "warnings":    warnings,
        "shifted":     shifted,
        "unscheduled": unscheduled
    }
