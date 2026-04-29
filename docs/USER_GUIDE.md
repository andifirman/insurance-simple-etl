# Project 001 — User Guide (Input & Terminal)

Dokumen ini menjelaskan:
1) Struktur file input Excel (sheet wajib + kolom yang dibutuhkan)
2) Input yang dimasukkan user lewat terminal (termasuk default)

## 1) File Input Excel

Sistem membaca file Excel yang dipilih user, lalu memproses menjadi:
- `data/processed/CF_Gen.xlsx`
- `data/processed/CSM_Gen.xlsx`

### Sheet yang wajib ada

#### A. `ValuationDate`
Digunakan untuk mengambil *Valuation Date*.

**Kolom minimal**
- Minimal harus ada **2 kolom** (kolom A & B).

**Aturan pembacaan Valuation Date**
- Sistem mencari baris yang kolom A-nya mengandung teks `VALUATION DATE` (case-insensitive), lalu mengambil tanggal dari kolom B.
- Jika tidak ketemu, sistem mengambil **nilai tanggal pertama** yang tidak kosong di kolom B.

**Catatan format tanggal**
- Excel date atau string tanggal (contoh `2023-12-31` atau `31/12/2023`).

---

#### B. `ICG`
Daftar unit yang akan digenerate.

**Kolom wajib**
- `ICG`

**Kolom opsional (disarankan, agar parsing lebih robust)**
- `Cohort`
- `Portfolio`
- `Contract`

Jika kolom opsional tidak ada, sistem akan mencoba *parse* dari string `ICG`.

---

#### C. `Earned`
Sumber data earned premium/commission per kombinasi `ICG` + `Incurred`.

**Kolom wajib**
- `ICG`
- `Incurred`
- `Earned_Premium`
- `Earned_Commission`

**Catatan horizon**
- Sistem akan memastikan horizon CF cukup panjang untuk menutup kebutuhan data di sheet `Earned`.
- Jika `Earned` mengandung periode lebih panjang dari End Date yang user input, sistem akan **auto-extend** horizon dan menampilkan pesan `[INFO] ...`.

---

#### D. `Expected_CF`
Sumber data expected premium/commission/acquisition per `ICG` + `Incurred`.

**Kolom wajib**
- `ICG`
- `Incurred`
- `Expected_Premium`
- `Expected_Commission`
- `Expected_Acquisition`

---

#### E. `Actual_CF`
Sumber data actual per `ICG` + `Incurred`.

**Kolom wajib**
- `ICG`
- `Incurred`
- `Actual_Premium`
- `Actual_Commission`
- `Actual_Acquisition`

---

#### F. `Assumptions`
Sumber ratio/asumsi per `ICG`.

**Kolom wajib**
- `ICG`
- `Loss_Ratio`
- `Risk_Adjustment_Ratio`
- `PME_Ratio`
- `ULAE_Ratio`
- `Premium_Refund_Ratio`
- `Cancellation per year (%)` *(atau `Cancellation_Ratio` pada template lama)*

**Kolom opsional (inflation berbasis assumptions)**
- `Inflation per year (%)` (nama header alternatif juga didukung)

**Catatan unit %**
- Jika kolom cancellation/inflation diisi `5` untuk 5%, sistem akan auto-scale menjadi `0.05`.

---

#### G. `Locked_in_Rate`
Sumber locked-in rate per `ICG`.

**Kolom wajib**
- `ICG`
- `Locked_in_Rate`

---

#### H. `Current_Rate`
Sumber current rate per `ICG`.

**Kolom wajib**
- `ICG`
- `Current_Rate`

---

### Sheet yang boleh ada (opsional)
- `Dropdown COB` (jika ada di template, sistem akan abaikan untuk kalkulasi)


## 2) Input User via Terminal

### 2.1 Pilih file input
- Sistem menampilkan file picker (window) untuk memilih Excel.
- File akan disalin ke `data/raw/`.

### 2.2 Pilih ICG yang mau digenerate
Setelah file dibaca, sistem menampilkan daftar `ICG` dari sheet `ICG`.

Kontrol (Windows Terminal):
- `↑/↓` pindah cursor
- `Space` checklist on/off
- `A` pilih semua
- `N` kosongkan pilihan
- `Enter` konfirmasi
- `Q` / `Esc` batal

Jika batal / tidak memilih apa pun, program berhenti.

### 2.3 Start Date
Prompt:
- `Masukkan Start Date (default YYYY-01-01, Enter untuk pakai default):`

Default Start Date:
- Diambil dari **tahun Valuation Date** → `01-Jan` tahun tersebut.

Dampak ke horizon:
- Data akan digenerate mulai dari **Q1** di tahun Start Date.

### 2.4 End Date
Prompt:
- `Masukkan End Date (format yyyy-mm-dd atau dd/mm/yyyy, Enter untuk pakai default):`

Default End Date:
- Diambil dari **tahun Valuation Date** → `31-Dec` tahun tersebut.

Catatan auto-extend:
- Jika kebutuhan horizon dari sheet `Earned` lebih panjang dari End Year yang user input, sistem akan **auto-extend** sampai cukup.


## Output
- `data/processed/CF_Gen.xlsx`
- `data/processed/CSM_Gen.xlsx`
