# 🌱 Arisan Berjaya

Web app arisan digital **PT. Berjaya Makmur Indonesia**, dibangun dengan Google Apps Script + Google Sheets.

- **Admin (laptop):** dashboard monitoring, grup, peserta, verifikasi setoran, galeri bukti, undian, kalender, laporan, pengaturan.
- **Peserta (HP):** login PIN, kirim setoran + foto struk, riwayat, unduh kwitansi PDF, foto profil.

## Struktur file

| File | Isi |
|---|---|
| `code.gs` | Backend: login PIN, CRUD, undian, upload Drive, kwitansi PDF |
| `setup.gs` | Inisialisasi database (Safe Migrate) + data dummy + folder Drive |
| `index.html` | Struktur halaman |
| `css.html` | Semua style |
| `js.html` | Semua logika frontend |
| `appsscript.json` | Manifest (timezone Asia/Jakarta, setting web app) |

## Pasang manual (tanpa clasp)

1. Buat Google Spreadsheet baru, lalu buka **Extensions → Apps Script**.
2. Buat file sesuai tabel di atas. File HTML diberi nama `index`, `css`, dan `js` (tanpa `.html`).
3. Buka **Project Settings**, lalu set **Time zone** ke `(GMT+07:00) Jakarta`.
4. Jalankan fungsi `setupDatabase` dan izinkan akses (Spreadsheet + Google Drive).
5. Pilih **Deploy → New deployment → Web app**, set *Execute as: Me* dan *Who has access: Anyone*.

Login dummy (dibuat otomatis kalau database masih kosong):

- Admin: `081200000000` / PIN `123456`. Segera ganti PIN setelah login pertama.
- Peserta: `081211110001` s/d `081211110010` / PIN `111111`.

> `setupDatabase` aman dijalankan berulang. Kolom baru ditambahkan di kanan, data lama tidak dihapus.

## Simpan ke GitHub

Butuh [Git for Windows](https://git-scm.com/download/win).

```bash
git init
git add .
git commit -m "Arisan Berjaya: versi awal"
git branch -M main
git remote add origin https://github.com/<username>/arisan-berjaya.git
git push -u origin main
```

Buat dulu repo **Private** bernama `arisan-berjaya` di GitHub, tanpa README, supaya push pertama tidak bentrok.

## Auto-deploy dengan GitHub Actions (opsional)

Workflow `.github/workflows/deploy.yml` mengirim kode ke Apps Script setiap kali ada push ke `main`. Sebelum secret diisi, workflow hanya dilewati dan tidak menghasilkan error.

1. Install [Node.js](https://nodejs.org) lalu jalankan `npm install -g @google/clasp@2.4.2`.
2. Aktifkan Apps Script API di https://script.google.com/home/usersettings.
3. Jalankan `clasp login`. Setelah itu file `C:\Users\<nama>\.clasprc.json` akan terbentuk.
4. Di GitHub, buka **Settings → Secrets and variables → Actions → New repository secret**, lalu isi:
   - `CLASPRC_JSON`: seluruh isi file `.clasprc.json`.
   - `SCRIPT_ID`: ID project Apps Script (di Project Settings → IDs).
   - `DEPLOYMENT_ID` (opsional): ID deployment web app (di Deploy → Manage deployments). Kalau diisi, URL web app otomatis memakai versi terbaru.

> ⚠️ `.clasprc.json` berisi token akses akun Google kamu. Jangan pernah meng-commit file ini. File ini sudah masuk `.gitignore`.

## Penyimpanan file di Google Drive

`setupDatabase` membuat folder `Arisan Berjaya` dengan 3 subfolder:

- `Foto Profil`
- `Struk Setoran`
- `Kwitansi`

File di folder-folder ini dibagikan sebagai *siapa saja yang punya link boleh melihat*, supaya gambar bisa tampil di aplikasi. Kalau akun Workspace kamu memblokir sharing publik, aplikasi otomatis memuat gambar lewat backend. Cara ini sedikit lebih lambat, tapi tetap jalan.
