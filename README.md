# 🧪 Mixing Schedule Planner

Aplikasi penjadwalan otomatis proses mixing berdasarkan jadwal filling.

## Cara Install & Jalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Alur Penggunaan

### 1. Tab Master Data (Setup sekali oleh Admin)

**Master Mixer:**
| Kolom | Keterangan |
|---|---|
| Mixer | Nama mixer (misal: Mixer A) |
| Kapasitas_kg | Kapasitas per batch dalam kg |
| Grup_Cleaning | Grup untuk menentukan cleaning (misal: Grup 1) |

**Master Produk:**
| Kolom | Keterangan |
|---|---|
| Kode_Produk | Kode unik produk |
| Nama_Produk | Nama produk |
| Grup_Cleaning | Grup cleaning produk |
| Kg_per_CS | Berat per CS dalam kg (misal: 12.5) |
| Mixer_Kompatibel | Mixer yang bisa dipakai, pisah koma (misal: Mixer A, Mixer B) |

### 2. Tab Input Planning (Planner setiap minggu)

Input per item:
- Kode produk (pilih dari dropdown)
- **Target dalam CS** → otomatis dikonversi ke kg
- Tanggal & shift filling
- Status: Urgent / Tidak Urgent

Atau upload bulk via Excel template.

### 3. Tab Jadwal Mixing

Klik **Generate Jadwal Mixing** → output otomatis:
- Jadwal per mixer, per shift
- Kolom: Tanggal, Shift, Mixer, Produk, Batches, Kapasitas Mixer, Total CS, Total kg
- 🔴 Merah = shift cleaning
- 🟡 Kuning = jadwal filling digeser (tidak urgent)
- Download ke Excel

## Aturan Penjadwalan

- Mixing harus selesai **minimal 1 shift sebelum** filling
- 1 shift = **2 batch** per mixer
- Ganti grup produk di mixer = **1 shift cleaning** (mixer lain tidak terpengaruh)
- **Urgent**: jadwal filling dikunci, mixing mundur menyesuaikan
- **Tidak Urgent**: jadwal filling bisa digeser dalam 1 minggu yang sama
