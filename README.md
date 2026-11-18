# SIP-RKAS V5 — Administrasi BOSP Reguler

Aplikasi administrasi transaksi dan arsip SPJ berbasis **Google Spreadsheet + Google Apps Script + React/Vite**. Google Spreadsheet adalah satu-satunya database utama; Google Drive digunakan untuk menyimpan PDF dan lampiran.

## Fitur V5
- Multi-user: Bendahara dan Kepala Sekolah dengan pembatasan keputusan persetujuan.
- Input transaksi satu kali beserta rincian barang; nomor dokumen otomatis.
- Pembuatan 5 dokumen PDF dasar: kuitansi, surat pesanan, pemeriksaan barang, BAST, dan rekap belanja.
- Paket ZIP per transaksi dan per kegiatan, lampiran bukti di Google Drive, checklist kelengkapan SPJ.
- Data penyedia, kegiatan/rekening, impor pagu CSV, rekap anggaran dan realisasi.
- Pencatatan pajak/penyesuaian manual dengan jenis, dasar, tarif input, nilai, dan referensi; **tidak menetapkan tarif pajak otomatis**.
- Alur persetujuan Kepala Sekolah: setujui/tolak, pembatalan transaksi secara soft-void (tidak menghapus riwayat).
- Penguncian periode/tahun anggaran; hanya Kepala Sekolah yang dapat menutup/membuka kembali.
- Audit log, riwayat persetujuan, backup salinan Spreadsheet ke Google Drive.
- Pengaturan identitas sekolah dan template dokumen.

## Struktur paket
- `gas/Code.gs` — backend Google Apps Script.
- `frontend/` — aplikasi React/Vite.
- `.env.example` — contoh URL Web App Apps Script.

## Instalasi
1. Buat Google Spreadsheet baru sebagai database.
2. Buka **Extensions → Apps Script**, salin isi `gas/Code.gs` ke file `Code.gs`.
3. Jalankan `setupSIPRKAS()` satu kali dari editor Apps Script dan berikan izin yang diminta.
4. Deploy → New deployment → Web app. Execute as: **Me**. Who has access: **Anyone** (atau kebijakan akses organisasi yang sesuai).
5. Salin URL `/exec` hasil deployment.
6. Pada folder `frontend`, salin `.env.example` menjadi `.env`, lalu isi `VITE_API_URL` dengan URL `/exec`.
7. Jalankan `npm install`, lalu `npm run build` untuk menghasilkan `dist`. Deploy folder `dist` ke Vercel/hosting statis. Set environment variable `VITE_API_URL` pada hosting.
8. Login awal: `kepsek` atau `bendahara`, password sementara `GantiSaya123!`. Segera ganti password dari menu Pengaturan.

## Catatan peningkatan dari V3
V5 menambahkan tabel `TAX_LEDGER`, `APPROVALS`, dan `PERIODS` secara otomatis saat `setupSIPRKAS()` dijalankan. Jalankan setup kembali setelah memasang V5 agar sheet baru dibuat. Setup mempertahankan data sheet yang sudah ada.

## Catatan penting sebelum digunakan resmi
- Paket ini **belum diuji integrasi pada akun Google sekolah Anda**. Jalankan build dan uji penerapan di salinan Spreadsheet terlebih dahulu sebelum data riil.
- Tarif, objek, kewajiban pemotongan/penyetoran pajak, format SPJ, dan dokumen wajib harus diverifikasi dengan juknis BOSP serta ketentuan perpajakan yang berlaku. Aplikasi tidak menggantikan ARKAS, sistem resmi pemerintah, pemeriksaan, atau persetujuan pejabat berwenang.
- PDF adalah template administrasi yang dapat disesuaikan. Dokumen hanya boleh ditandatangani setelah transaksi, barang, pemeriksaan, dan serah terima benar-benar terjadi. Jangan membuat bukti fiktif.
- Persetujuan dalam aplikasi adalah jejak workflow internal, bukan tanda tangan elektronik tersertifikasi.
- Password memakai SHA-256 tanpa salt khusus; untuk penggunaan produksi, gunakan kebijakan akun yang kuat, batasi akses Web App/folder Drive, dan pertimbangkan autentikasi Google Workspace.
- Penguncian periode menggunakan tahun dari tanggal transaksi. Penguncian tidak mengubah data historis.
- Pajak/penyesuaian hanya dicatat saat transaksi dibuat; tidak ada kalkulasi tarif otomatis dan tidak ada pelaporan pajak elektronik.
- V5 tidak mengintegrasikan API ARKAS atau sistem perpajakan eksternal.

## Kompatibilitas
Apps Script V8. Frontend React + Vite. Zona waktu backend `Asia/Jakarta`.
