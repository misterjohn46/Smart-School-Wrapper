# Smart School Wrapper

Penghubung (Client Bridge) yang menghubungkan Google Spreadsheet Anda dengan **CoreSystem Library** Smart School.

## Apa itu Wrapper?

`Kode.js` adalah script penghubung yang harus di-copy ke Apps Script di spreadsheet Anda. Script ini menyediakan fungsi-fungsi (seperti `doGet`, `getToken`, dsb.) yang memanggil method di **CoreSystem Library**.

```
┌─────────────────┐     ┌──────────────┐     ┌─────────────────┐
│  Google Sheet   │ ──▶ │  Kode.js     │ ──▶ │  CoreSystem     │
│  (Spreadsheet)  │     │  (Wrapper)   │     │  Library        │
└─────────────────┘     └──────────────┘     └─────────────────┘
```

## Setup Awal (Pertama Kali)

### 1. Buka Apps Script

Buka spreadsheet Anda, lalu klik **Extensions ▶ Apps Script**.

### 2. Hapus kode default

Hapus semua isi file `Code.gs` yang muncul.

### 3. Copy Kode.js

Copy seluruh isi dari [Kode.js](Kode.js) ke file `Code.gs` di Apps Script.

### 4. Tambahkan Library CoreSystem

1. Di Apps Script editor, klik **+** di samping **Libraries**
2. Masukkan Script ID CoreSystem (didapat dari administrator Smart School)
3. Klik **Look up**
4. Pilih versi terbaru (atau `0` untuk development)
5. Ganti **Identifier** menjadi `CoreSystem`
6. Klik **Add**

### 5. Atur OAuth Scopes

Buka **Project Settings ▶ Show "appsscript.json" manifest** dan pastikan scopes ini ada:

```json
{
  "oauthScopes": [
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/userinfo.email"
  ]
}
```

### 6. Deploy Web App (untuk API endpoints)

1. Klik **Deploy ▶ New deployment**
2. Pilih type: **Web App**
3. Execute as: **User deploying**
4. Who has access: **Anyone**
5. Klik **Deploy**
6. **Copy URL yang muncul** — ini adalah URL API Anda

### 7. Setup Sheet "Pengaturan" (Opsional)

Buat sheet bernama `Pengaturan` dengan format:

| KEY | VALUE |
|-----|-------|
| NAMA_SEKOLAH | SMA Negeri 1 ... |
| NPSN | 12345678 |

Template WA bisa ditambahkan di sini (KEY: `WA_TEMPLATE_MASUK`, `WA_TEMPLATE_PULANG`, `WA_TEMPLATE_GURU_MASUK`, `WA_TEMPLATE_GURU_PULANG`).

---

## Cara Update Wrapper (Untuk Klien)

Saat ada update terbaru, klien bisa meng-update script mereka dengan 3 cara. Pilih yang paling mudah:

---

### ⭐ Metode 1: Web Updater (Paling Mudah — untuk non-teknisi)

Gunakan halaman web updater. Tidak perlu buka Apps Script editor!

1. Buka **[update.html](https://misterjohn46.github.io/Smart-School-Wrapper/update.html)** di browser
2. Masukkan **Script ID** project Apps Script Anda
   > Cara dapat Script ID: buka Apps Script → **Project Settings** → scroll ke **IDs** → copy **Script ID**
3. Klik **Connect Google Account** → login dengan akun Google yang punya akses ke Apps Script
4. Klik **🚀 Push Kode.js**
5. **Selesai!** Refresh editor Apps Script untuk melihat kode terbaru.

> **Keuntungan:** Tidak perlu copy-paste, tidak perlu buka editor. Cukup 3 klik.

---

### ⚡ Metode 2: Auto-Update dari Apps Script

Jalankan fungsi update langsung dari dalam Apps Script editor:

1. Buka spreadsheet Anda → **Extensions ▶ Apps Script**
2. Dari dropdown fungsi (dekat tombol Run), pilih **`updateWrapper`**
3. Klik **Run** → beri izin jika diminta saat popup muncul
4. Tunggu sampai log muncul `✅ updateWrapper BERHASIL!`
5. **Refresh** halaman editor (F5) untuk melihat kode terbaru
6. **PENTING:** Deploy ulang Web App:
   - **Deploy ▶ Manage deployments**
   - Klik ikon ✏️ (edit) pada deployment aktif
   - Pilih **Version → New version**
   - Klik **Deploy**

> **Prasyarat:**
> 1. `appsscript.json` harus punya scope `https://www.googleapis.com/auth/script.projects` (sudah otomatis jika menggunakan `appsscript.json` dari repo ini).
> 2. **Apps Script API harus diaktifkan di GCP project.** Jika error "Apps Script API has not been used", buka URL yang muncul di log → klik **Enable** → tunggu 1-2 menit → coba lagi.
> 3. Jika tetap gagal, gunakan **Metode 1 (Web Updater)** — tidak perlu enable API manual.


#### Cek Versi Dulu

Sebelum update, cek apakah wrapper Anda sudah versi terbaru:

1. Pilih fungsi **`checkWrapperVersion`** dari dropdown
2. Klik **Run**
3. Buka **View ▶ Logs**. Jika perbedaan >5 baris, jalankan update.

---

### 📋 Metode 3: Copy Manual (Fallback)

Gunakan jika kedua metode di atas tidak bisa:

1. Buka [Kode.js di GitHub](https://raw.githubusercontent.com/misterjohn46/Smart-School-Wrapper/main/Kode.js)
2. **Ctrl+A** (Select All) lalu **Ctrl+C** (Copy)
3. Buka Apps Script editor spreadsheet Anda (**Extensions ▶ Apps Script**)
4. **Ctrl+A** lalu **Delete** — hapus SEMUA isi `Code.gs` yang lama
5. **Ctrl+V** (Paste) kode baru
6. Klik ikon **💾 Save**
7. Deploy ulang Web App (lihat langkah 6 di Metode 2)

---

### 🛠 Metode 4: Via clasp (Untuk Developer)

```bash
git clone https://github.com/misterjohn46/Smart-School-Wrapper.git
cd Smart-School-Wrapper
npx @google/clasp login
echo '{"scriptId":"SCRIPT_ID_ANDA"}' > .clasp.json
npx @google/clasp push
```

---

### Update Library CoreSystem

Jika library CoreSystem di-update (biasanya diinformasikan oleh admin Smart School):

1. Buka Apps Script editor
2. Klik **Libraries** di panel kiri
3. Pilih library `CoreSystem`
4. Ganti **Version** ke versi terbaru
5. Klik **Save**

---

## Ringkasan Metode Update

| Metode | Cocok Untuk | Kesulitan | Waktu |
|--------|-------------|-----------|-------|
| 1. Web Updater | Semua klien (non-teknisi) | ⭐ Sangat Mudah | ~1 menit |
| 2. Auto-Update | Klien yang sudah setup awal | ⭐⭐ Mudah | ~2 menit |
| 3. Copy Manual | Fallback / pertama kali | ⭐⭐⭐ Sedang | ~5 menit |
| 4. clasp | Developer | ⭐⭐⭐⭐ Mahir | ~3 menit |

> **Rekomendasi:** Gunakan selalu **Metode 1: Web Updater** — cukup buka link, login, klik Push.

---

## Troubleshooting

### Error: "CoreSystem is not defined"
- Pastikan library CoreSystem sudah di-add dengan **Identifier** `CoreSystem`
- Cek di **Project Settings** apakah library muncul

### Error: "Authorization required"
- Jalankan fungsi apapun dari editor sekali untuk trigger popup OAuth
- Atau klik **Run ▶ Review Permissions** dan izinkan semua scope

### Web App URL tidak merespon
- Pastikan deployment masih aktif: **Deploy ▶ Manage deployments**
- Jika ada perubahan kode, buat **New version** di deployment

---

## Structure Repo

```
Smart-School-Wrapper/
├── Kode.js          # Bridge utama (satu-satunya file yang dibutuhkan)
└── README.md        # Dokumentasi ini
```

