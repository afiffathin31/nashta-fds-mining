# NashTa FDS — Coal Mining & Cash Forensics Edition

**Produk:** Fraud Detection System (NashTa × JendelaTax) — Mining & Cash Forensics Edition  
**Tipe Arsitektur:** Forensic & Behavioral Intelligence Overlay (POJK No. 12/2024)  
**Dataset Utama:** 68 Transaksi Bersih Buku Kas Bank Mandiri (Akun 1000001) & 8 Aturan Forensik Tambang (`dummy.xlsx`).

---

## 🎯 Fokus Kasus & Modul Forensik
1. **Buku Kas & Bank Mandiri Forensics (Akun 1000001):**
   - Rekonsiliasi 68 transaksi kas Bank Mandiri periode Januari 20x8.
   - Deteksi Double Transfer Jasa Hauling PT. MAS Rp 980 Juta (`001/BIN/I/20x8`).
   - Audit rekening perantara / suspense PT. AADC Rp 7,415 Miliar (`1199001`).
   - 8 Aturan Forensik & Kriteria Audit (*Auditor Notes*).
2. **Kontraktor & Vendor Hauling Tambang:**
   - Deteksi Benturan Kepentingan (*Conflict of Interest*) PT. Perkasa Kuat Sejahtera (Perusahaan Pribadi Direktur Utama).
   - OCR Invoice Checksum & ELA Tamper Analysis `07/INT-JJJ/XII/2023`.
   - Kolusi IP Address lelang tender hauling batubara (`192.168.1.45`).
3. **Pajak & DJP Compliance (JendelaTax):**
   - 3-Way Equalization: Buku Kas Mandiri $\leftrightarrow$ SPT Masa $\leftrightarrow$ Server DJP.
   - PPh 22 Pembelian Batubara Rp 510 Juta (`014/BOU`).
   - PPh 23 Sewa Alat & Jasa Perbaikan Site (`012 & 013/BOU`).
   - Audit Restitusi PPh Badan Rp 4,9 Miliar (`016/BOU`).
4. **Kontrol Internal & Matriks SoD 5x5:**
   - Pemisahan wewenang kasir vs otorisasi transfer bank.
   - Log override limit approval pinjaman Rp 5 Miliar.
   - Analisis *Rubber Stamping* 35 PO dalam 2.4 detik.
5. **Logistik Tambang & Jembatan Timbang Km 107:**
   - Rekap ritase penimbangan batubara 2.006 MT (PT. MAS), 410 MT (PT. Ninja), 452 MT (Koperasi IBL).
   - Deteksi Phantom Hauling (GPS void).
   - Selisih stockopname 1.450 MT di Stockpile ISP Titan Km 107.
   - Reject batubara high-ash (> 18%).
6. **Digital Dossier Workspace (Case AL-2024-001):**
   - Multi-Hop Money Flow SVG Network Graph interaktif.
   - Kronologi detik-demi-detik aliran dana kas.
   - Digital Evidence Vault ber-hash SHA-256.

---

## 🚀 Cara Menjalankan di Lokal
1. Jalankan web server statis:
   ```cmd
   start.bat
   ```
   Atau via Python:
   ```bash
   python -m http.server 5175
   ```
2. Buka browser pada `http://localhost:5175`.
