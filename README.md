# 📊 ETL Dashboard Pendapatan - User Guide

Automated ETL Pipeline untuk Generate Laporan Pendapatan Bulanan (Januari - Desember)

---

## 🚀 Cara Menjalankan

### **Opsi 1: Command Line (Rekomendasi untuk Production)**

```powershell
python main.py
```

**Konfigurasi**: Edit file `main.py` bagian `CONFIG` (line 16-28):

```python
CONFIG = {
    "INPUT_FILE": "12 Lampiran Pendapatan Desember 2025.xlsx",  # File lampiran pendapatan bulan ini
    "FILE_REKAP": "01 251110 Validasi Kode Produk Rekap New.xlsx",  # File master validasi produk
    "FILE_PELANGGAN": "12. BANGER-PELANGGAN-12122025.xlsx",  # File data pelanggan
    "FILE_OPT": "12. BANGER-OPT-12122025.xlsx",  # File data OPT
    "TEMPLATE_FILE": "11 BANGER NOVEMBER (2).xlsx",  # File template bulan lalu
    "OUTPUT_FILE": "12 BANGER DESEMBER ayolah2.xlsx",  # File output
    "BULAN_LALU": "November",  # Bulan dari template
    "BULAN_INI": "Desember",   # Bulan yang diproses
}
```

---

### **Opsi 2: Streamlit Web App (User-Friendly)**

```powershell
streamlit run app.py
```

Buka browser di `http://localhost:8501`, lalu:

1. **Upload 5 file** di sidebar (Template, Lampiran, Rekap, Pelanggan, OPT)
2. **Pilih bulan** dari dropdown (Bulan Lalu & Bulan Ini)
3. **Klik "GENERATE LAPORAN"**
4. **Download hasil** Excel dari browser

---

## 📦 Installation & Requirements

### **1. Install Python 3.12+**

Download dari: https://www.python.org/downloads/

### **2. Install Dependencies**

```powershell
pip install pandas openpyxl streamlit
```

**Package yang dibutuhkan**:

- `pandas` - Data manipulation & ETL
- `openpyxl` - Excel file handling (.xlsx)
- `streamlit` - Web UI (optional, hanya untuk `app.py`)

### **3. Verifikasi Installation**

```powershell
python --version  # Should show Python 3.12+
pip list | Select-String "pandas|openpyxl|streamlit"
```

---

## 📋 File Input Requirements

### **1. Lampiran Pendapatan (INPUT_FILE)**

**Format**: `.xlsx` atau `.xlsb`

**Sheet yang dibaca**: Sheet dengan nama mengandung **"Konsol"** (case-insensitive)

**Struktur Header** (Row 1-2):

- **Row 1**: Nama produk lengkap (e.g., "Dark Fiber", "Clear Channel")
- **Row 2**: Kode produk (angka: 120, 121, 122, ...) + kolom customer

**Kolom STATIS yang WAJIB ada** (nama harus **PERSIS SAMA**):

| Kolom Statis             | Alternatif Nama                  | Wajib?      | Keterangan                          |
| ------------------------ | -------------------------------- | ----------- | ----------------------------------- |
| `Customer No`            | `Customer NO`, `Customer Number` | ✅ Wajib    | Nomor customer (akan dinormalisasi) |
| `Customer Name`          | -                                | ⚠️ Optional | Nama customer (bisa kosong)         |
| `120`, `121`, `122`, ... | (kode produk angka)              | ✅ Wajib    | Kolom revenue per produk            |

**⚠️ CRITICAL: Normalisasi Data Sebelum Upload**

Sebelum run script, **WAJIB cek & fix manual** di Excel:

1. **Customer Number format aneh**:
   - ❌ Salah: `2000XXXXX.` (ada titik di belakang)
   - ✅ Benar: `2000021345` (pure angka/text, tanpa titik/koma)
2. **Customer Number blank**:
   - ❌ Salah: `(blank)` atau kosong
   - ✅ Benar: `888888888` atau `999999999` (placeholder konsisten)

**Cara cek di Excel**:

```
1. Filter kolom "Customer No" → cari "(blank)" atau yang ada titik
2. Replace semua titik (.) dengan kosong (Ctrl+H)
3. Isi blank cells dengan placeholder (e.g., 888888888)
4. Save & close Excel sebelum run script
```

---

### **2. File Rekap (FILE_REKAP)**

**Format**: `.xlsx`

**Sheet yang dibaca**: `"ALL PRODUCT PDF"` (header di **row 3**)

**Kolom STATIS yang WAJIB ada**:

| Kolom Statis                  | Wajib?      | Keterangan                               |
| ----------------------------- | ----------- | ---------------------------------------- |
| `ICON+ Product`               | ✅ Wajib    | Kode produk master (e.g., 756, 719, 731) |
| `Product Portofolio Segmen 1` | ⚠️ Optional | Nama produk level 1                      |
| `Product Portofolio Segmen 2` | ⚠️ Optional | Nama produk level 2                      |
| `Product Portofolio Segmen 3` | ⚠️ Optional | Nama produk level 3                      |
| `SEGMEN`                      | ⚠️ Optional | Segmen produk                            |

**Fallback nama produk**: Script cari nama produk dari kolom segmen (prioritas: Segmen 1 → Segmen 2 → Segmen 3 → SEGMEN)

---

### **3. Data Pelanggan (FILE_PELANGGAN)**

**Format**: `.xlsx`

**Sheet yang dibaca**: Sheet pertama (index 0)

**Kolom STATIS yang WAJIB ada** (nama **CASE-SENSITIVE**):

| Kolom Statis                                | Wajib?   | Keterangan                              |
| ------------------------------------------- | -------- | --------------------------------------- |
| `idPerusahaan`                              | ✅ Wajib | ID perusahaan                           |
| `idCustomerSap`                             | ✅ Wajib | ID customer SAP                         |
| `idPelanggan`                               | ✅ Wajib | ID pelanggan                            |
| `namaPerusahaan`                            | ✅ Wajib | Nama perusahaan                         |
| `nomorKontrak`                              | ✅ Wajib | Nomor kontrak                           |
| `kodeMasterProduk`                          | ✅ Wajib | Kode produk (untuk sorting & filtering) |
| `namaMasterProduk`                          | ✅ Wajib | Nama produk                             |
| `carryOverJanuari`                          | ✅ Wajib | Carry Over bulan Januari                |
| `carryOverFebruari` - `carryOverDesember`   | ✅ Wajib | Carry Over bulan Feb-Des (12 kolom)     |
| `newRevenueJanuari`                         | ✅ Wajib | New Revenue bulan Januari               |
| `newRevenueFebruari` - `newRevenueDesember` | ✅ Wajib | New Revenue bulan Feb-Des (12 kolom)    |

**Total kolom minimum**: 41 kolom (5 kolom info + 12 Carry Over + 12 New Revenue + 12 tambahan)

---

### **4. Data OPT (FILE_OPT)**

**Format**: `.xlsx`

**Sheet yang dibaca**: Sheet pertama (index 0)

**Kolom STATIS**: **SAMA PERSIS** dengan Data Pelanggan (41 kolom yang sama)

---

### **5. Template (TEMPLATE_FILE)**

**Format**: `.xlsx`

**Sheet yang WAJIB ada**:

#### **Sheet: Dashboard**

**Struktur Row**:

- **Row 1**: Header besar (bebas, tidak disentuh script)
- **Row 2**: Header kolom (Kode Produk, Nama Produk, Januari, ..., Desember)
- **Row 3-63**: Data produk (61 produk + 2 cadangan)
- **Row 64**: **GRAND TOTAL** (HARDCODED - WAJIB ADA!)

**Kolom STATIS (hardcoded di script)**:

| Kolom                         | Letter | Index   | Keterangan               |
| ----------------------------- | ------ | ------- | ------------------------ |
| Kode Produk                   | A      | 1       | Kode produk              |
| Nama Produk                   | B      | 2       | Nama produk              |
| Kumulatif Januari             | J      | 10      | Kumulatif bulan Januari  |
| Kumulatif Februari - November | K - T  | 11 - 20 | Kumulatif bulan Feb-Nov  |
| Kumulatif Desember            | U      | 21      | Kumulatif bulan Desember |
| Sisa                          | V      | 22      | Sisa prognosa            |
| SA Januari                    | W      | 23      | Stand Alone Januari      |
| SA Februari - November        | X - AG | 24 - 33 | Stand Alone Feb-Nov      |
| SA Desember                   | AH     | 34      | Stand Alone Desember     |

**⚠️ CRITICAL**: Row 64 (Grand Total) **HARUS ADA** di template!

---

#### **Sheet: Summary**

**Struktur**:

- **B2**: "Target Desember" (text, STATIS)
- **C2**: "Realisasi {Bulan}" (text, DINAMIS - diupdate script)
- **B3**: **Angka target manual** (e.g., `326470130576`) - **ISI MANUAL DI TEMPLATE!**
- **C3**: Formula (DINAMIS - diupdate script → `=Dashboard!U64`)

**⚠️ CRITICAL**: Cell **B3** (Target Desember) **WAJIB diisi manual** di template!

---

#### **Sheet: Data Pelanggan**

**Row 2 Config** (WAJIB ADA):

- **AU (col 47)**: "Bulan Berjalan" (label, STATIS)
- **AV (col 48)**: Angka bulan (1-12, DINAMIS - diupdate script)

**Kolom Carry Over & New Revenue**:

- **AY (col 51)**: Carry Over Januari
- **AZ - BJ (col 52-62)**: Carry Over Feb-Des
- **BK (col 63)**: New Revenue Januari
- **BL - BV (col 64-74)**: New Revenue Feb-Des

---

#### **Sheet: Data OPT**

**Row 2 Config** (WAJIB ADA):

- **CE (col 83)**: "Bulan Berjalan" (label, STATIS)
- **CF (col 84)**: Angka bulan (1-12, DINAMIS - diupdate script)

---

## 🔄 Workflow ETL

### **FASE 1: EXTRACT**

1. Baca master produk dari `FILE_REKAP` (sheet "ALL PRODUCT PDF")
2. Baca data raw dari `INPUT_FILE` (sheet "Konsol")
3. **Normalisasi Customer Number** (remove dots/commas - **SUDAH MANUAL FIX**)
4. Filter data:
   - Drop rows tanpa Customer Number
   - Drop rows "Grand Total"
   - Drop pivot summary (Digital Platform, PLN Group, dll)

### **FASE 2: TRANSFORM**

1. **Unpivot data** (kode produk jadi rows)
2. **Mapping nama produk** (kode → nama dari master)
3. **Filter & sort** by 61 kode produk custom order:
   ```
   756, 719, 731, 714, 464, 715, 737, 750, 725, 740, ...
   ```
4. **Aggregate** revenue per produk

### **FASE 3: LOAD**

1. **Copy template** → output file
2. **Create/update sheet** `Realisasi {Bulan}`
3. **Update Dashboard** dengan logika dinamis:
   - **Januari** (bulan_index=1): Copy prognosa baseline ke semua bulan
   - **Feb-Nov** (bulan_index=2-11):
     - Row 3: Prognosa = Rata-rata
     - Row 4-63: Prognosa = SUMIF Data Pelanggan
   - **Desember** (bulan_index=12): Prognosa jadi Realisasi (SUMIF)
4. **Update Summary**:
   - C2: "Realisasi {Bulan}"
   - C3: `=Dashboard!{Kumulatif}{GrandTotal}`
5. **Update Data Pelanggan & OPT**:
   - Set Bulan Berjalan (1-12)
   - Replace data dengan file baru (sorted & filtered)

---

## ⚙️ Konfigurasi Kode Produk

**File**: `main.py` line 33-42

**61 Kode Produk** (custom sort order):

```python
CUSTOM_KODE_PRODUK_ORDER = [
    756, 719, 731, 714, 464, 715, 737, 750, 725, 740, 747, 727, 739, 749, 748,
    721, 743, 745, 744, 735, 717, 738, 692, 673, 698, 693, 699, 697, 696, 694,
    695, 723, 746, 736, 713, 730, 728, 729, 755, 741, 724, 726, 757, 758, 732,
    273, 425, 435, 436, 734, 733, 722, 274, 708, 709, 710, 764, 768, 766, 767,
    487
]
```

**Fungsi**:

- Sorting data (Realisasi, Data Pelanggan, Data OPT)
- Filtering (drop kode yang tidak ada di list)

**⚠️ Update list ini** jika ada perubahan produk!

---

## 🐛 Troubleshooting

### **Error: "Kolom wajib tidak ditemukan"**

**Penyebab**: Nama kolom di file input **tidak sesuai** dengan yang diharapkan script.

**Solusi**:

1. Cek nama kolom di file Excel (case-sensitive!)
2. Rename kolom sesuai tabel "Kolom STATIS" di atas
3. Contoh fix:
   - ❌ `customer no` → ✅ `Customer No`
   - ❌ `KodeMasterProduk` → ✅ `kodeMasterProduk` (case-sensitive!)

---

### **Error: "Sheet 'Konsol' tidak ditemukan"**

**Penyebab**: Lampiran Pendapatan tidak punya sheet dengan nama mengandung "Konsol".

**Solusi**:

1. Cek nama sheet di Excel (case-insensitive, tapi harus ada kata "Konsol")
2. Rename sheet jadi `"Konsol"` atau `"Data Konsol"`

---

### **Warning: "Dropped XXX rows (kode produk tidak ada di list 61 kode)"**

**Penyebab**: Ada kode produk di data yang **tidak ada** di list `CUSTOM_KODE_PRODUK_ORDER`.

**Solusi**:

1. **Jika produk valid**: Tambahkan kode ke list (line 33-42 di `main.py`)
2. **Jika produk invalid**: Biarkan (akan di-drop otomatis)

---

### **Total Revenue Tidak Match**

**Penyebab**: Ada Customer Number dengan format aneh (titik, koma) yang bikin SUMIF ga match.

**Solusi**:

1. **Manual fix** di Excel SEBELUM run script:
   - Cari & replace titik `.` dengan kosong
   - Isi blank cells dengan `888888888`
2. Save Excel
3. Run script lagi

---

### **Summary B3 (Target Desember) Kosong**

**Penyebab**: Template tidak punya angka target di cell B3.

**Solusi**:

1. Buka template Excel
2. Isi cell **Summary!B3** dengan angka target (e.g., `326470130576`)
3. Save template
4. Script **TIDAK AKAN** overwrite cell ini (dibiarin statis)

---

## 📝 Best Practices

### **1. Backup Template**

Sebelum run script, **copy template** jadi backup:

```powershell
Copy-Item "11 BANGER NOVEMBER.xlsx" "11 BANGER NOVEMBER (BACKUP).xlsx"
```

### **2. Manual Validation**

Setelah run script, **cek manual**:

- [ ] Row 64 (Grand Total) punya formula SUM vertikal
- [ ] Summary C3 reference ke `Dashboard!U64` (atau kolom kumulatif bulan ini)
- [ ] Tab `Realisasi {Bulan}` ada & sorted by 61 kode produk
- [ ] Data Pelanggan & OPT punya Bulan Berjalan (1-12)

### **3. Normalisasi Data Input**

Sebelum upload, **WAJIB fix**:

- [ ] Customer Number tanpa titik/koma (pure angka/text)
- [ ] Blank Customer Number diisi placeholder (888888888)
- [ ] Nama kolom sesuai list "Kolom STATIS"

### **4. Update Kode Produk**

Jika ada perubahan produk:

1. Update list `CUSTOM_KODE_PRODUK_ORDER` (line 33-42)
2. Update template (tambah/hapus row produk di Dashboard)
3. Test dengan data sample sebelum production

---

## 📞 Support

**Developer**: ETL Team  
**File Konfigurasi**: `main.py` (line 16-42)  
**UI Web App**: `app.py` (Streamlit)

**Log Output**: Check terminal untuk detail proses ETL:

```
🚀 [1/3] EXTRACT: Mapping Master Data...
⚙️ [2/3] TRANSFORM: Unpivoting & Matching...
💾 [3/3] LOAD: Update Dashboard dengan Logika Dinamis S-R...
✅ BERHASIL! Dashboard + Summary + Data Pelanggan + OPT dinamis untuk DESEMBER!
```

---



## ✅ Checklist Sebelum Run

- [ ] Python 3.12+ installed
- [ ] Dependencies installed (`pandas`, `openpyxl`, `streamlit`)
- [ ] **Customer Number di Lampiran sudah dinormalisasi** (no dots/commas, no blanks)
- [ ] Template punya Row 64 (Grand Total)
- [ ] Template Summary B3 diisi angka target manual
- [ ] Nama kolom di semua file sesuai "Kolom STATIS"
- [ ] Config `BULAN_LALU` & `BULAN_INI` sudah benar
- [ ] File paths di `CONFIG` sudah benar

---

**Last Updated**: January 9, 2026  
**Version**: 1.0 (Dynamic for all 12 months)
