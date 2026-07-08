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
2. Masukkan Script ID: `1TEjwz_Jrr16mgSpNmjTs8uZuREWC_MNSk_x2a9odI3DThUJy09easBiU`
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

## Cara Update Wrapper

Saat ada update terbaru dari repository ini:

### Metode 1: Copy Manual (Rekomended)

1. Buka file [Kode.js](Kode.js) di repo ini
2. Klik **Raw** (atau copy seluruh isi)
3. Buka Apps Script di spreadsheet Anda
4. **Hapus seluruh isi** `Code.gs` yang lama
5. **Paste** kode baru
6. Klik **Save** (💾)
7. **Deploy ulang** Web App jika ada perubahan: **Deploy ▶ Manage deployments ▶ Edit ▶ Version ▶ New version ▶ Deploy**

### Metode 2: Via clasp (untuk developer)

```bash
# Clone repo wrapper
git clone https://github.com/misterjohn46/Smart-School-Wrapper.git
cd Smart-School-Wrapper

# Login clasp
npx @google/clasp login

# Buat .clasp.json dengan script ID Anda
echo '{"scriptId":"SCRIPT_ID_ANDA"}' > .clasp.json

# Push ke Apps Script
npx @google/clasp push
```

### Update Library Version

Jika CoreSystem library di-update:

1. Buka Apps Script editor
2. Klik **Libraries**
3. Pilih `CoreSystem`
4. Ganti versi ke yang terbaru
5. Save

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

