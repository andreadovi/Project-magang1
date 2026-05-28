import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from scheduler import generate_mixing_schedule
from pivot import build_pivot, pivot_to_excel

st.set_page_config(page_title="Mixing Scheduler", page_icon="🧪", layout="wide")
st.title("🧪 Mixing Schedule Planner")

DAYS_ID = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

# ─── Session State ────────────────────────────────────────────────────────────
if "master_mixer"  not in st.session_state:
    st.session_state.master_mixer  = pd.DataFrame(columns=["Mixer", "Kapasitas_kg", "Grup_Cleaning"])
if "master_produk" not in st.session_state:
    st.session_state.master_produk = pd.DataFrame(columns=["Kode_Produk", "Nama_Produk", "Grup_Cleaning", "Kg_per_CS", "Resting_Days", "Mixer_Kompatibel"])
if "filling_plan"  not in st.session_state:
    st.session_state.filling_plan  = pd.DataFrame()

tab1, tab2, tab3 = st.tabs(["⚙️ Master Data", "📋 Input Planning", "📅 Jadwal Mixing"])

# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 — MASTER DATA
# ═════════════════════════════════════════════════════════════════════════════
with tab1:
    st.header("Master Data")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🔧 Master Mixer")
        tmpl_mixer = pd.DataFrame({
            "Mixer":         ["Mixer A", "Mixer B", "Mixer C"],
            "Kapasitas_kg":  [500, 800, 300],
            "Grup_Cleaning": ["Grup 1", "Grup 1", "Grup 2"]
        })
        buf = io.BytesIO(); tmpl_mixer.to_excel(buf, index=False)
        st.download_button("📥 Download Template Mixer", buf.getvalue(),
                           "template_master_mixer.xlsx", use_container_width=True)

        up_mixer = st.file_uploader("Upload Master Mixer", type=["xlsx", "csv"], key="mixer_upload")
        if up_mixer:
            df  = pd.read_excel(up_mixer) if up_mixer.name.endswith("xlsx") else pd.read_csv(up_mixer)
            req = {"Mixer", "Kapasitas_kg", "Grup_Cleaning"}
            if req.issubset(df.columns):
                st.session_state.master_mixer = df
                st.success(f"✅ {len(df)} mixer berhasil diupload!")
            else:
                st.error(f"❌ Kolom harus: {req}")

        if not st.session_state.master_mixer.empty:
            st.dataframe(st.session_state.master_mixer, use_container_width=True, hide_index=True)

    with col2:
        st.subheader("📦 Master Produk")
        st.caption("**Resting_Days**: `2` = perlu didiamkan 2 hari, `0` = tidak. **Mixer_Kompatibel**: pisah koma.")
        tmpl_produk = pd.DataFrame({
            "Kode_Produk":      ["P001", "P002", "P003"],
            "Nama_Produk":      ["Produk Alpha", "Produk Beta", "Produk Gamma"],
            "Grup_Cleaning":    ["Grup 1", "Grup 1", "Grup 2"],
            "Kg_per_CS":        [12.5, 8.0, 10.0],
            "Resting_Days":     [0, 2, 0],
            "Mixer_Kompatibel": ["Mixer A, Mixer B", "Mixer B", "Mixer C"]
        })
        buf2 = io.BytesIO(); tmpl_produk.to_excel(buf2, index=False)
        st.download_button("📥 Download Template Produk", buf2.getvalue(),
                           "template_master_produk.xlsx", use_container_width=True)

        up_produk = st.file_uploader("Upload Master Produk", type=["xlsx", "csv"], key="produk_upload")
        if up_produk:
            df  = pd.read_excel(up_produk) if up_produk.name.endswith("xlsx") else pd.read_csv(up_produk)
            req = {"Kode_Produk", "Nama_Produk", "Grup_Cleaning", "Kg_per_CS", "Resting_Days", "Mixer_Kompatibel"}
            if req.issubset(df.columns):
                st.session_state.master_produk = df
                st.success(f"✅ {len(df)} produk berhasil diupload!")
            else:
                st.error(f"❌ Kolom harus: {req}")

        if not st.session_state.master_produk.empty:
            st.dataframe(st.session_state.master_produk, use_container_width=True, hide_index=True)

# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 — INPUT PLANNING
# ═════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header("Input Planning Filling")

    if st.session_state.master_produk.empty:
        st.warning("⚠️ Upload Master Produk dulu di tab Master Data.")
    else:
        # ── Week picker ───────────────────────────────────────────────────────
        st.subheader("📆 Pilih Minggu Filling")
        filling_week = st.date_input(
            "Pilih tanggal mana saja dalam minggu filling",
            value=datetime.today(),
            key="filling_week_picker"
        )
        # Get Monday of selected week
        week_monday = filling_week - timedelta(days=filling_week.weekday())
        week_dates  = [week_monday + timedelta(days=i) for i in range(7)]
        week_labels = [f"{DAYS_ID[d.weekday()]} {d.strftime('%d/%m')}" for d in week_dates]

        # Shift columns
        shift_cols = []
        shift_meta = []  # (date_str, shift_num)
        for d, label in zip(week_dates, week_labels):
            for s in [1, 2, 3]:
                shift_cols.append(f"{label} S{s}")
                shift_meta.append((d.strftime("%Y-%m-%d"), s))

        st.caption(f"Minggu: **{week_monday.strftime('%d %b')} — {week_dates[-1].strftime('%d %b %Y')}**")

        # ── Build editable grid ───────────────────────────────────────────────
        st.subheader("📋 Tabel Planning (isi jumlah CS)")
        st.caption("Kosongkan sel jika tidak ada filling. Centang **Urgent** per produk.")

        produk_df = st.session_state.master_produk

        # Init grid per week
        grid_key = f"grid_{week_monday.strftime('%Y%m%d')}"
        if grid_key not in st.session_state:
            init_data = {
                "Urgent":       [False] * len(produk_df),
                "Kode_Produk":  list(produk_df["Kode_Produk"]),
                "Nama_Produk":  list(produk_df["Nama_Produk"]),
            }
            for col in shift_cols:
                init_data[col] = [None] * len(produk_df)
            st.session_state[grid_key] = pd.DataFrame(init_data)

        edited_df = st.data_editor(
            st.session_state[grid_key],
            use_container_width=True,
            hide_index=True,
            key=f"editor_{grid_key}",
            column_config={
                "Urgent":      st.column_config.CheckboxColumn("🚨 Urgent", default=False),
                "Kode_Produk": st.column_config.TextColumn("Kode Produk", disabled=True),
                "Nama_Produk": st.column_config.TextColumn("Nama Produk", disabled=True),
                **{col: st.column_config.NumberColumn(col, min_value=0, step=1)
                   for col in shift_cols}
            }
        )
        # Persist edits without triggering rerun
        st.session_state[grid_key] = edited_df

        if st.button("💾 Simpan Planning", type="primary", use_container_width=True):
            # Convert grid to long format filling_plan
            rows = []
            for _, row in edited_df.iterrows():
                kode   = row["Kode_Produk"]
                nama   = row["Nama_Produk"]
                urgent = "Urgent" if row["Urgent"] else "Tidak Urgent"
                for col, (date_str, shift_num) in zip(shift_cols, shift_meta):
                    val = row[col]
                    if pd.notna(val) and val is not None and float(val) > 0:
                        rows.append({
                            "Kode_Produk":     kode,
                            "Nama_Produk":     nama,
                            "Target_CS":       float(val),
                            "Tanggal_Filling": date_str,
                            "Shift_Filling":   shift_num,
                            "Urgent":          urgent
                        })

            if rows:
                st.session_state.filling_plan = pd.DataFrame(rows)
                st.success(f"✅ {len(rows)} item planning tersimpan!")
            else:
                st.warning("⚠️ Tidak ada data yang diisi.")

        # Show saved plan summary
        if not st.session_state.filling_plan.empty:
            with st.expander("📋 Lihat Planning Tersimpan"):
                st.dataframe(st.session_state.filling_plan, use_container_width=True, hide_index=True)
                st.caption(f"Total: {len(st.session_state.filling_plan)} item")

# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 — JADWAL MIXING
# ═════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header("Jadwal Mixing Otomatis")

    ready = (not st.session_state.master_mixer.empty and
             not st.session_state.master_produk.empty and
             not st.session_state.filling_plan.empty)

    if not ready:
        st.warning("⚠️ Lengkapi Master Mixer, Master Produk, dan Input Planning terlebih dahulu.")
    else:
        # ── Mixer schedule range: Jumat before filling week, 10 days ─────────
        st.subheader("📆 Range Jadwal Mixing")

        # Default: Friday before the filling week
        filling_plan_df = st.session_state.filling_plan
        min_fill_date   = pd.to_datetime(filling_plan_df["Tanggal_Filling"]).min()
        # Find Friday of previous week
        days_to_friday  = (min_fill_date.weekday() - 4) % 7
        default_start   = (min_fill_date - timedelta(days=days_to_friday + 7)).date()
        default_end     = (default_start + timedelta(days=9))

        c1, c2 = st.columns(2)
        with c1:
            mix_start = st.date_input("Dari tanggal (Jumat)", value=default_start)
        with c2:
            mix_end   = st.date_input("Sampai tanggal", value=default_end)

        date_range = []
        d = mix_start
        while d <= mix_end:
            date_range.append(d.strftime("%Y-%m-%d"))
            d += timedelta(days=1)
        st.caption(f"Range mixing: **{mix_start.strftime('%d %b')} — {mix_end.strftime('%d %b %Y')}** ({len(date_range)} hari)")

        if st.button("⚡ Generate Jadwal Mixing", type="primary", use_container_width=True):
            with st.spinner("Menjadwalkan mixing..."):
                result = generate_mixing_schedule(
                    st.session_state.master_mixer,
                    st.session_state.master_produk,
                    st.session_state.filling_plan
                )
            st.session_state.schedule_result = result

        if "schedule_result" in st.session_state:
            result      = st.session_state.schedule_result
            schedule_df = result["schedule"]
            warnings    = result["warnings"]
            shifted     = result["shifted"]
            unscheduled = result["unscheduled"]

            if warnings:
                for w in warnings:
                    st.warning(w)

            if shifted:
                st.subheader("🔀 Jadwal Filling Digeser (Tidak Urgent)")
                st.dataframe(pd.DataFrame(shifted).style.map(
                    lambda _: "background-color: #fff3cd"),
                    use_container_width=True, hide_index=True)

            if unscheduled:
                st.subheader("❌ Tidak Bisa Dijadwalkan")
                for u in unscheduled:
                    st.error(u)

            if not schedule_df.empty:
                st.subheader("📅 Jadwal Mixing")

                pivot_df, meta = build_pivot(
                    schedule_df,
                    st.session_state.master_mixer,
                    st.session_state.master_produk,
                    date_range
                )

                if not pivot_df.empty:
                    display_df = pivot_df.copy()
                    rename_map = {c: c.replace("\n", " ") for c in display_df.columns}
                    display_df = display_df.rename(columns=rename_map)
                    col_display = [c.replace("\n", " ") for c in meta["col_labels"]]

                    def style_pivot(row):
                        styles = [""] * len(row)
                        mixer  = row.get("Mixer", "")
                        kode   = row.get("Kode_Produk", "")
                        cols   = list(row.index)
                        for i, label in enumerate(col_display):
                            d, s = meta["col_keys"][i]
                            if label not in cols:
                                continue
                            pos = cols.index(label)
                            if (mixer, kode, d, s) in meta["cleaning_cells"]:
                                styles[pos] = "background-color: #BDD7EE"
                            elif (mixer, kode, d, s) in meta["resting_cells"]:
                                styles[pos] = "background-color: #FFE699"
                        return styles

                    st.dataframe(
                        display_df.style.apply(style_pivot, axis=1),
                        use_container_width=True, hide_index=True
                    )
                    st.caption("🔵 Biru = Cleaning")

                    excel_data = pivot_to_excel(pivot_df, meta, st.session_state.master_mixer)
                    st.download_button(
                        "📥 Download Excel",
                        excel_data,
                        "jadwal_mixing.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                with st.expander("📋 Detail Jadwal (Raw)"):
                    st.dataframe(
                        schedule_df.drop(columns=["Cleaning"], errors="ignore"),
                        use_container_width=True, hide_index=True
                    )

                with st.expander("🔍 Debug: Pivot Rows"):
                    if not pivot_df.empty:
                        st.write("Baris di pivot:", pivot_df[["Mixer","Kode_Produk","Nama_Produk"]].to_dict("records"))
                    else:
                        st.write("Pivot kosong!")
                    st.write("Unscheduled:", result["unscheduled"])
