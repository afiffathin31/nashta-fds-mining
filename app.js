window.chartInstances = {}; const chartInstances = window.chartInstances;
// =========================================================================
// NASHTA FDS - COAL MINING & CASH MANAGEMENT FORENSICS EDITION
// DATASET: C:\Users\Apip\Downloads\dummy.xlsx (Dataset Bersih Bank Mandiri 1000001)
// =========================================================================

// --- STATE MANAGEMENT ---
const state = {
  currentView: 'financial',
  currentSubView: {
    financial: 'gl',
    procurement: 'matrix',
    tax: 'equalization',
    approvals: 'sod_matrix',
    operational: 'timbangan'
  },
  
  // 68 TRANSAKSI BERSIH BUKU KAS BANK MANDIRI DARI DUMMY.XLSX
  bankMandiriLedger: [
  {
    "row": 3,
    "date": "04 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "001/-/BOU/I/20x8",
    "description": "Biaya OperaLLKonal",
    "debit": 0.0,
    "kredit": 600.0,
    "balance": 257.0,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 4,
    "date": "04 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "003/-/BOU/I/20x8",
    "description": "Biaya Admin",
    "debit": 0.0,
    "kredit": 250.0,
    "balance": 7.0,
    "flag": "HIGH: Biaya Admin Tidak Wajar Rp 250 Jt",
    "severity": "warning"
  },
  {
    "row": 5,
    "date": "04 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "003/-/BOU/I/20x8",
    "description": "Final uncrushed coal 410,460 MT @ 320.000 dari stock ROM PT. NInja ke Stockpile ISP Titan Km 107 - F-183/INV/TG-TAN2/XII/2023",
    "debit": 0.0,
    "kredit": 146.0,
    "balance": -139.0,
    "flag": "OPS: Stock ROM Stockpile Titan Km 107",
    "severity": "normal"
  },
  {
    "row": 6,
    "date": "04 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "002/-/BOU/I/20x8",
    "description": "Biaya Admin",
    "debit": 0.0,
    "kredit": 2.5,
    "balance": -141.5,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 7,
    "date": "04 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "002/-/BOU/I/20x8",
    "description": "Stockopname  - 003/PO/JJJ-SCF/XII/2023 - PI 20040123000697",
    "debit": 0.0,
    "kredit": 215.0,
    "balance": -356.5,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 8,
    "date": "04 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "004/-/BOU/I/20x8",
    "description": "Iuran BPJS Kesehatan HO Bulan Jan 2024",
    "debit": 0.0,
    "kredit": 1.2,
    "balance": -357.7,
    "flag": "HR AUDIT: Iuran BPJS",
    "severity": "normal"
  },
  {
    "row": 9,
    "date": "04 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "005/-/BOU/I/20x8",
    "description": "(Koperasi IBL) Jasa transportasi angkutan batubara dari Kalibata Ke stockpile PT. MAS 452,17 MT jarak 24,4 km  @ 1.600 - 007/INV/KKK-JJJ/XII/23",
    "debit": 0.0,
    "kredit": 19.0,
    "balance": -376.7,
    "flag": "OPS: Jasa Hauling Batubara",
    "severity": "normal"
  },
  {
    "row": 10,
    "date": "05 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "007/-/BOU/I/20x8",
    "description": "Jasa Transportir Coal Hauling to IPD Alam4 PT. Mas 60 MT Jarak58,1 km @ 1.200 - 09/INT/JJJ/XII/2023",
    "debit": 0.0,
    "kredit": 2.5,
    "balance": -379.2,
    "flag": "OPS: Jasa Hauling Batubara",
    "severity": "normal"
  },
  {
    "row": 11,
    "date": "05 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "001/-/BIN/I/20x8",
    "description": "DOUBLE TF -  Jasa transportir coal hauling to IPD Alam4 PT. MAS 2.006,03 MT jarak 25,1 km @ 1.600  periode tagihan 28 Nov sd 10 Desember 2023 - 07/INT-JJJ/XII/2023",
    "debit": 980.0,
    "kredit": 0.0,
    "balance": 600.8,
    "flag": "CRITICAL: Double Transfer Rp 980 Jt (001/BIN)",
    "severity": "danger"
  },
  {
    "row": 12,
    "date": "05 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "006/-/BOU/I/20x8",
    "description": "Iuran Bpjs ketenagakerjaan periode Januari 2024 (gaji desember 2023) HO",
    "debit": 0.0,
    "kredit": 5.1,
    "balance": 595.6999999999999,
    "flag": "HR AUDIT: Gaji GL Rp 112 Jt vs HRD Rp 80 Jt (Selisih Rp 32 Jt)",
    "severity": "warning"
  },
  {
    "row": 13,
    "date": "08 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "008/-/BOU/I/20x8",
    "description": "(PT. Perkasa Kuat Sejahtera) Reimbursement royalti LLK No. 150/LLK/JJJ-Kalibata/I/2024 8.700 MT - 004/Kalibata/INV/I/2024",
    "debit": 0.0,
    "kredit": 49.0,
    "balance": 546.6999999999999,
    "flag": "CONFLICT OF INTEREST: Perusahaan Pribadi Dirut",
    "severity": "danger"
  },
  {
    "row": 14,
    "date": "08 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "010/-/BOU/I/20x8",
    "description": "Pinjaman Sementara (Royalti Provisi Batubara 5% PT. Kalibata 8.700 MT) LLK 148/LLK/JJJ-Kalibata/I/2024",
    "debit": 0.0,
    "kredit": 70.0,
    "balance": 476.69999999999993,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 15,
    "date": "08 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "009/-/BOU/I/20x8",
    "description": "Biaya Trafficman dan pemberLLKhan jalan aspal desa manggul periode 1-30 Desember 2023 - 006-SRQ-ENG-KDM-0124",
    "debit": 0.0,
    "kredit": 6.0,
    "balance": 470.69999999999993,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 16,
    "date": "08 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "002/-/BIN/I/20x8",
    "description": "Pinjaman Sementara - PT. AADC",
    "debit": 700.0,
    "kredit": 0.0,
    "balance": 1170.6999999999998,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 17,
    "date": "09 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "003/-/BIN/I/20x8",
    "description": "Pinjaman Sementara  - PT. AADC",
    "debit": 715.0,
    "kredit": 0.0,
    "balance": 1885.6999999999998,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 18,
    "date": "09 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "011/-/BOU/I/20x8",
    "description": "Pinjaman Sementara (Royalti ProviLLKonal Batubara 8% PT. Kalibata 8.700 MT) LLK 523",
    "debit": 0.0,
    "kredit": 500.0,
    "balance": 1385.6999999999998,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 19,
    "date": "10 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "014/-/BOU/I/20x8",
    "description": "Pph 22 Masa Des 2023",
    "debit": 0.0,
    "kredit": 510.0,
    "balance": 875.6999999999998,
    "flag": "TAX: Setoran PPh 22 Batubara Rp 510 Juta",
    "severity": "normal"
  },
  {
    "row": 20,
    "date": "10 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "005/-/BIN/I/20x8",
    "description": "(PT. Kuda Mas Paint) Penerimaan pendapatan atas final jasa pengangkutan batubara 11,58 MT @75.000 Inv No. 219/INV/JJJ-KDM-28F/XII/2023",
    "debit": 4.3,
    "kredit": 0.0,
    "balance": 879.9999999999998,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 21,
    "date": "10 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "006/-/BIN/I/20x8",
    "description": "(PT. Kuda Mas Paint) Penerimaan pendapatan atas final jasa pengangkutan batubara 27,99 MT @75.000 Inv No. 218/INV/JJJ-KDM-27F/XII/2023",
    "debit": 7.8,
    "kredit": 0.0,
    "balance": 887.7999999999997,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 22,
    "date": "10 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "007/-/BIN/I/20x8",
    "description": "(PT. Kuda Mas Paint) Penerimaan pendapatan atas final jasa pengangkutan batubara 1,01 MT @75.000 Inv No. 217/INV/JJJ-KDM-25F/XII/2023",
    "debit": 0.83,
    "kredit": 0.0,
    "balance": 888.6299999999998,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 23,
    "date": "10 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "004/-/BIN/I/20x8",
    "description": "(PT. Jon Coal) Penerimaan pendapatan atas final penjualan batubara 1,01 MT @375.000  Inv No. 220/INV/JJJ-JJJ-1F/XII/2023",
    "debit": 0.14,
    "kredit": 0.0,
    "balance": 888.7699999999998,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 24,
    "date": "10 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "016/-/BOU/I/20x8",
    "description": "Penerimaan restitusi PPh Badan",
    "debit": 4900.0,
    "kredit": 0.0,
    "balance": 5788.7699999999995,
    "flag": "TAX: Restitusi PPh Badan Rp 4.9 Miliar",
    "severity": "normal"
  },
  {
    "row": 25,
    "date": "10 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "012/-/BOU/I/20x8",
    "description": "Pph 23 rental masa des 2023",
    "debit": 0.0,
    "kredit": 26.0,
    "balance": 5762.7699999999995,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 26,
    "date": "10 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "013/-/BOU/I/20x8",
    "description": "Pph 23 service masa des 2023",
    "debit": 0.0,
    "kredit": 0.23,
    "balance": 5762.54,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 27,
    "date": "10 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "015/-/BOU/I/20x8",
    "description": "Pph 4 (2) Masa Des 2023",
    "debit": 0.0,
    "kredit": 1.0,
    "balance": 5761.54,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 28,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "025/-/BOU/I/20x8",
    "description": "(PT. Pahala Kehidupan) Jasa angkutan KDM- MAS 985 MT @ 40.160 periode 18-26 Des 2023 - 001/INV/DT/ESA-JJJ/I/2024",
    "debit": 0.0,
    "kredit": 59.0,
    "balance": 5702.54,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 29,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "022/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Final pembelian batubara 1.01 MT @ 345.000 tujuan franco baturaja - F 047/KDM/INV/XI/2023",
    "debit": 0.0,
    "kredit": 6.0,
    "balance": 5696.54,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 30,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "017/-/BOU/I/20x8",
    "description": "(PT. Perkasa Kuat Sejahtera) Reimbursement royalti LLK No. 151/LLK/JJJ-Kalibata/I/2024 8.700 MT - 005/Kalibata/INV/I/2024",
    "debit": 0.0,
    "kredit": 6.0,
    "balance": 5690.54,
    "flag": "CONFLICT OF INTEREST: Perusahaan Pribadi Dirut",
    "severity": "danger"
  },
  {
    "row": 31,
    "date": "11 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "008/-/BIN/I/20x8",
    "description": "Pinjaman Sementara - PT. AADC",
    "debit": 5000.0,
    "kredit": 0.0,
    "balance": 10690.54,
    "flag": "HIGH: Pinjaman Sementara Rp 5.0 Miliar (PT. AADC)",
    "severity": "danger"
  },
  {
    "row": 32,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "018/-/BOU/I/20x8",
    "description": "(PT. Pelangi Dimatamu) Reimbursement royalti LLK 123 JJJ 1.x00 MT - 100/NInja/INV/XII/2023",
    "debit": 0.0,
    "kredit": 52.0,
    "balance": 10638.54,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 33,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "019/-/BOU/I/20x8",
    "description": "(PT. Pelangi Dimatamu) Reimbursement royalti LLK 144 - 092/NInja/INV/XII/2023",
    "debit": 0.0,
    "kredit": 222.0,
    "balance": 10416.54,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 34,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "021/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Final pembelian batubara 27,99 MT @ 325.000 tujuan franco ulak pandan - F 049/KDM/INV/XII/2023",
    "debit": 0.0,
    "kredit": 22.4,
    "balance": 10394.140000000001,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 35,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "020/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Final pembelian batubara 11,58 MT @ 305.000 tujuan franco stocpile PT. MAS - F 050/KDM/INV/XII/2023",
    "debit": 0.0,
    "kredit": 17.0,
    "balance": 10377.140000000001,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 36,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "023/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Reimbursement royalti Final LLK 324 5.015,46 - 053/KDM/INV/XII/2023",
    "debit": 0.0,
    "kredit": 7.9,
    "balance": 10369.240000000002,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 37,
    "date": "11 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "024/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Reimbursement royalti TI 124 JJJ 1.000 MT - 051/KDM/INV/XII/2023",
    "debit": 0.0,
    "kredit": 34.04,
    "balance": 10335.2,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 38,
    "date": "11 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "026 A/-/BOU/I/20x8",
    "description": "Biaya Admin TF ke PT. AAA",
    "debit": 0.0,
    "kredit": 0.5,
    "balance": 10334.7,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 39,
    "date": "12 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "009/-/BIN/I/20x8",
    "description": "(PT. Jon Coal) Penerimaan atas Penjualan Batubara FOT 1.000 MT @375.000 Inv No. 002/INV/JJJ-JJJ-1/I/2024",
    "debit": 410.0,
    "kredit": 0.0,
    "balance": 10744.7,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 40,
    "date": "12 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "009 A/-/BIN/I/20x8",
    "description": "Kelebihan terima atas Penjualan Batubara FOT Inv No. 002/INV/JJJ-JJJ-1/I/2024",
    "debit": 3.8,
    "kredit": 0.0,
    "balance": 10748.5,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 41,
    "date": "12 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "026/-/BOU/I/20x8",
    "description": "(PT. Merdana Indi) Final Jasa angkutan batubara 524 MT jarak 24,4 km @ 1.600 periode 14 des - 26 Des 2023 - 17/INV-FINAL/AAA/XII/2023",
    "debit": 0.0,
    "kredit": 21.0,
    "balance": 10727.5,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 42,
    "date": "15 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "027/-/BOU/I/20x8",
    "description": "(PT. Pelangi Dimatamu) Final pembelian batubara 0,13 MT @ 345.000 Tujuan Franco PT Semen Baturaja - F 048/NInja/INV/IX/2023",
    "debit": 0.0,
    "kredit": 0.4,
    "balance": 10727.1,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 43,
    "date": "15 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "010/-/BIN/I/20x8",
    "description": "Penerimaan atas Pinjaman Sementara (Royalti ProviLLKonal Batubara 8% PT. Kalibata 8.700 MT) LLK 148/LLK/JJJ-Kalibata/I/2024",
    "debit": 685.0,
    "kredit": 0.0,
    "balance": 11412.1,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 44,
    "date": "16 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "011 A/-/BIN/I/20x8",
    "description": "Kelebihan terima atas final penjualan batubara 27,99 MT @320.000  Inv No. 001/INV/JJJ-SMJ-10F/I/2024",
    "debit": 0.2,
    "kredit": 0.0,
    "balance": 11412.300000000001,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 45,
    "date": "16 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "011/-/BIN/I/20x8",
    "description": "(PT. Kuda Mas) Penerimaan pendapatan atas final penjualan batubara 27,99 MT @320.000  Inv No. 001/INV/JJJ--SMJ-10F/I/2024",
    "debit": 15.0,
    "kredit": 0.0,
    "balance": 11427.300000000001,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 46,
    "date": "16 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "028/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Pembayaran Royalti LLK No. 153/LLK/JJJ-KDM/I/2024 - 001/KDM/INV/I/2024",
    "debit": 0.0,
    "kredit": 154.0,
    "balance": 11273.300000000001,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 47,
    "date": "16 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "029/-/BOU/I/20x8",
    "description": "(PT. Perkasa Kuat Sejahtera) Pembayaran Royalti LLK No. 152/LLK/JJJ-Kalibata/I/2024 - 007/Kalibata/INV/I/2024",
    "debit": 0.0,
    "kredit": 342.0,
    "balance": 10931.300000000001,
    "flag": "CONFLICT OF INTEREST: Perusahaan Pribadi Dirut",
    "severity": "danger"
  },
  {
    "row": 48,
    "date": "16 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "030/-/BOU/I/20x8",
    "description": "Pembelian mesin  ABC 1unit",
    "debit": 0.0,
    "kredit": 1554.0,
    "balance": 9377.300000000001,
    "flag": "ASSET AUDIT: Pengadaan Mesin ABC Rp 1.55 M (Underspek)",
    "severity": "danger"
  },
  {
    "row": 49,
    "date": "16 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "012/-/BIN/I/20x8",
    "description": "Penerimaan atas Pinjaman Sementara (Royalti ProviLLKonal Batubara 8% PT. Kalibata 8.700 MT) LLK 523",
    "debit": 742.0,
    "kredit": 0.0,
    "balance": 10119.300000000001,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 50,
    "date": "16 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "031/-/BOU/I/20x8",
    "description": "Pembayaran Sewa Kost 2 Bln 17 Jan s.d 16 Mar 2024 a/n Bintang Lemon - 00005/FPD/HRD-HO/I/2024",
    "debit": 0.0,
    "kredit": 3.0,
    "balance": 10116.300000000001,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 51,
    "date": "16 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "032/-/BOU/I/20x8",
    "description": "Biaya Admin",
    "debit": 0.0,
    "kredit": 0.02,
    "balance": 10116.28,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 52,
    "date": "16 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "032/-/BOU/I/20x8",
    "description": "CloLLKng Stock inventory cargo JJJ dan Legal Certificate Surveyor Independent Tah",
    "debit": 0.0,
    "kredit": 45.0,
    "balance": 10071.28,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 53,
    "date": "18 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "014/-/BIN/I/20x8",
    "description": "Penerimaan atas Pinjaman Sementara  LLK 124",
    "debit": 54.0,
    "kredit": 0.0,
    "balance": 10125.28,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 54,
    "date": "18 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "015/-/BIN/I/20x8",
    "description": "Penerimaan atas Pinjaman Sementara (Royalti Provisional Batubara 15% PT. NInja 1.500 MT) LLK 123",
    "debit": 56.0,
    "kredit": 0.0,
    "balance": 10181.28,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 55,
    "date": "18 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "013/-/BIN/I/20x8",
    "description": "Penerimaan atas Pinjaman Sementara (Royalti Provisional Batubara 1% PT. KDM 8.300 MT) LLK 399",
    "debit": 850.0,
    "kredit": 0.0,
    "balance": 11031.28,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 56,
    "date": "22 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "033/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Pembelian Batubara 1.000 MT @380.000 Tujuan Franco Baturaja - 003/KDM/INV/I/2024",
    "debit": 0.0,
    "kredit": 582.0,
    "balance": 10449.28,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 57,
    "date": "22 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "033/-/BOU/I/20x8.",
    "description": "Lebih Bayar atas Pembelian Batubara 1.000 MT @380.000 Tujuan Franco Baturaja - 003/KDM/INV/I/2024 - PT. KDM",
    "debit": 0.0,
    "kredit": 3.85,
    "balance": 10445.43,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 58,
    "date": "26 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "016/-/BIN/I/20x8",
    "description": "(PT. Sugih Kaya Sugih Sentosa) Penerimaan atas Final Penjualan Batubara FCA Merapi Inv No 003 s.d 008 Jan 2024 BMSS",
    "debit": 2238.0,
    "kredit": 0.0,
    "balance": 12683.43,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 59,
    "date": "26 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "034/-/BOU/I/20x8",
    "description": "Pengembalian Pinjaman Sementara",
    "debit": 0.0,
    "kredit": 2500.0,
    "balance": 10183.43,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 60,
    "date": "29 Jan 20x8",
    "source": "Penerimaan Penjualan",
    "voucherNo": "017/-/BIN/I/20x8",
    "description": "(PT. Sugih Kaya Sugih Sentosa) Penerimaan atas Penjualan Batubara FCA Merapi 7.500 MT @456.400 Inv No. 010/INV/JJJ-BMSS-1/I/2024",
    "debit": 3748.0,
    "kredit": 0.0,
    "balance": 13931.43,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 61,
    "date": "30 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "035/-/BOU/I/20x8",
    "description": "Gaji Bulan Jan 2024",
    "debit": 0.0,
    "kredit": 112.0,
    "balance": 13819.43,
    "flag": "HR AUDIT: Gaji GL Rp 112 Jt vs HRD Rp 80 Jt (Selisih Rp 32 Jt)",
    "severity": "warning"
  },
  {
    "row": 62,
    "date": "31 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "040/-/BOU/I/20x8",
    "description": "(PT. Kuda Mas Paint) Pembelian Batubara >>  Tujuan Franco Merapi IPD MAS - 004/KDM/INV/I/2024",
    "debit": 0.0,
    "kredit": 2614.0,
    "balance": 11205.43,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 63,
    "date": "31 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "042/-/BOU/I/20x8",
    "description": "rounding",
    "debit": 0.0,
    "kredit": 7.4,
    "balance": 11198.03,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 64,
    "date": "31 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "042/-/BOU/I/20x8",
    "description": "Biaya Admin Bln Jan dan Pajak",
    "debit": 0.0,
    "kredit": 350.0,
    "balance": 10848.03,
    "flag": "HIGH: Biaya Admin Tidak Wajar Rp 250 Jt",
    "severity": "warning"
  },
  {
    "row": 65,
    "date": "31 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "018/-/BIN/I/20x8",
    "description": "Bunga Bank Bln Jan",
    "debit": 800.0,
    "kredit": 0.0,
    "balance": 11648.03,
    "flag": "Normal",
    "severity": "normal"
  },
  {
    "row": 66,
    "date": "31 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "041/-/BOU/I/20x8",
    "description": "Pelunasan Pinjaman Sementara  (2.450.000.000-2.000.000.000=450.000.000)",
    "debit": 0.0,
    "kredit": 50.0,
    "balance": 11598.03,
    "flag": "WARNING: Pinjaman Sementara Tanpa BAP",
    "severity": "warning"
  },
  {
    "row": 67,
    "date": "31 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "036/-/BOU/I/20x8",
    "description": "(PT. Perkasa Kuat Sejahtera) Pembayaran Royalti LLK No 3/LLK/JJJ-Kalibata/I/2024 - 003/Kalibata/INV/I/2024",
    "debit": 0.0,
    "kredit": 48.0,
    "balance": 11550.03,
    "flag": "CONFLICT OF INTEREST: Perusahaan Pribadi Dirut",
    "severity": "danger"
  },
  {
    "row": 68,
    "date": "31 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "037/-/BOU/I/20x8",
    "description": "(PT. Perkasa Kuat Sejahtera) Reimbursment royalti LLK No. 7/LLK/JJJ-Kalibata/XII/2023 - 101F/Kalibata/INV/I/2024",
    "debit": 0.0,
    "kredit": 13.0,
    "balance": 11537.03,
    "flag": "CONFLICT OF INTEREST: Perusahaan Pribadi Dirut",
    "severity": "danger"
  },
  {
    "row": 69,
    "date": "31 Jan 20x8",
    "source": "Pembayaran Pembelian",
    "voucherNo": "038/-/BOU/I/20x8",
    "description": "(PT. Perkasa Kuat Sejahtera) Reimbursment royalti LLK No.5R/LLK/JJJ-Kalibata/XI/2023 - 096F/Kalibata/INV/I/2024",
    "debit": 0.0,
    "kredit": 10.0,
    "balance": 11527.03,
    "flag": "CONFLICT OF INTEREST: Perusahaan Pribadi Dirut",
    "severity": "danger"
  },
  {
    "row": 70,
    "date": "31 Jan 20x8",
    "source": "Bukti Jurnal",
    "voucherNo": "039/-/BOU/I/20x8",
    "description": "Pengembalian kelebihan pembayaran Inv No. 002/INV/JJJ-JJJ-1/I/2024 - 001/CN/JJJ-JJJ/I/2024",
    "debit": 0.0,
    "kredit": 3.7,
    "balance": 11523.33,
    "flag": "Normal",
    "severity": "normal"
  }
],

  // KONTRAKTOR & VENDOR LIST
  vendors: [
    { name: 'PT. Perkasa Kuat Sejahtera', npwp: '31.229.182.1-018.000', spend: 'Rp 55.000.000', flags: 'CONFLICT OF INTEREST: Perusahaan Pribadi Direktur Utama', score: '96/100', mitigation: 'Wajib Audit Khusus', category: 'Royalty & Landowner' },
    { name: 'PT. MAS (Mitra Alam Sentosa)', npwp: '01.882.910.4-012.000', spend: 'Rp 3.840.000.000', flags: 'CRITICAL: Double Transfer Rp 980 Juta (001/BIN)', score: '98/100', mitigation: 'Payment Hold Aktif', category: 'Jasa Coal Hauling' },
    { name: 'PT. AADC (Anugerah Alam Duta Coal)', npwp: '02.441.819.3-021.000', spend: 'Rp 7.415.000.000', flags: 'HIGH: Pinjaman Sementara Gantung Rp 5.0 Miliar', score: '92/100', mitigation: 'Investigasi Direksi', category: 'Trading & Pembiayaan' },
    { name: 'PT. Pelangi Dimatamu (PT. Ninja)', npwp: '01.992.341.2-024.000', spend: 'Rp 420.000.000', flags: 'Stock ROM Km 107 Uncrushed Coal & Royalti 15%', score: '74/100', mitigation: 'Cross-Check Surveyor', category: 'Stockpile ROM Provider' },
    { name: 'PT. Kuda Mas Paint', npwp: '03.118.291.5-014.000', spend: 'Rp 40.500.000', flags: 'Split Invoicing (Inv 217, 218, 219 batubara franco)', score: '65/100', mitigation: 'Monitoring Kontrak', category: 'Supplier & Hauling' },
    { name: 'PT. Pahala Kehidupan', npwp: '02.910.481.6-031.000', spend: 'Rp 59.000.000', flags: 'Jasa Angkutan KDM-MAS 985 MT @ 40.160', score: '45/100', mitigation: 'Wajar / Normal', category: 'Jasa Transportir' },
    { name: 'Koperasi IBL', npwp: '01.554.891.2-019.000', spend: 'Rp 19.000.000', flags: 'Transportasi 452.17 MT Jarak 24.4 km', score: '30/100', mitigation: 'Wajar / Terverifikasi', category: 'Koperasi Warga Lokal' }
  ],

  // ANTREAN KASUS INVESTIGASI
  cases: [
    { id: 'AL-2024-001', title: 'Indikasi Double Transfer Jasa Hauling Batubara PT. MAS (Rp 980 Juta) & Rekening Siluman', module: 'Procurement & Cash Bank', entity: 'PT. MAS (Mitra Alam Sentosa)', amount: 'Rp 980.000.000', severity: 'High', status: 'In Progress', investigator: 'Budi Santoso, CFE' },
    { id: 'AL-2024-002', title: 'Arus Dana Pinjaman Sementara Berulang PT. AADC (Total Rp 7.415 Miliar) Tanpa BAP Direksi', module: 'Financial & UEBA', entity: 'PT. AADC (Trading Partner)', amount: 'Rp 5.000.000.000', severity: 'High', status: 'Triage Pending', investigator: 'Siti Rahma, CISA' },
    { id: 'AL-2024-003', title: 'Benturan Kepentingan PT. Perkasa Kuat Sejahtera (Perusahaan Pribadi Dirut)', module: 'Vendor & Royalty', entity: 'Direktur Utama / PT. Perkasa Kuat', amount: 'Rp 55.000.000', severity: 'Critical', status: 'Audit Khusus', investigator: 'Budi Santoso, CFE' },
    { id: 'AL-2024-004', title: 'Pengadaan Mesin ABC Rp 1.55 Miliar Underspek & Beda Merk (Voucher 030/BOU)', module: 'Fixed Assets Audit', entity: 'Supplier Mesin ABC', amount: 'Rp 1.554.000.000', severity: 'Critical', status: 'Asset Inspection', investigator: 'Doni Pratama, CFE' },
    { id: 'AL-2024-005', title: 'Selisih Gaji GL Rp 112 Jt vs Data Riil HRD Rp 80 Jt (Voucher 035/BOU)', module: 'Payroll Audit', entity: 'Finance Payroll', amount: 'Rp 32.000.000', severity: 'High', status: 'HR Cross-Check', investigator: 'Sri Wahyuni, BKP' }
  ]
};

// --- DOM CONTENT LOADED ENTRYPOINT ---
document.addEventListener('DOMContentLoaded', () => {
  initNavigation();
  initAllCharts();
  initLiveSearch();
});

// --- LIVE SEARCH IN TOPBAR ---
function initLiveSearch() {
  const input = document.getElementById('global-search-input');
  if (input) {
    input.addEventListener('input', (e) => {
      const q = e.target.value.toLowerCase().trim();
      renderBankMandiriTable(q);
    });
  }
}

// --- RENDER BANK MANDIRI TABLE WITH LIVE SEARCH / FILTER ---
function renderBankMandiriTable(query = '', flagFilter = 'all') {
  const table = document.getElementById('table-fin-gl-table');
  if (!table) return;

  let filtered = state.bankMandiriLedger;
  if (query) {
    filtered = filtered.filter(t => 
      t.description.toLowerCase().includes(query) ||
      t.voucherNo.toLowerCase().includes(query) ||
      t.source.toLowerCase().includes(query) ||
      t.date.toLowerCase().includes(query) ||
      t.flag.toLowerCase().includes(query)
    );
  }

  if (flagFilter !== 'all') {
    filtered = filtered.filter(t => t.severity === flagFilter);
  }

  table.innerHTML = `
    <thead>
      <tr>
        <th style="width: 45px;">No</th>
        <th style="width: 100px;">Tanggal</th>
        <th style="width: 140px;">Sumber & Voucher</th>
        <th>Uraian Transaksi Batubara & Rekening</th>
        <th class="text-right" style="width: 130px;">Debit (Kas Masuk)</th>
        <th class="text-right" style="width: 130px;">Kredit (Kas Keluar)</th>
        <th class="text-right" style="width: 130px;">Saldo Kas Mandiri</th>
        <th style="width: 160px;">Flag Sistem FDS</th>
        <th class="text-right" style="width: 90px;">Aksi</th>
      </tr>
    </thead>
    <tbody>
      ${filtered.map((t, idx) => {
        const debitStr = t.debit > 0 ? `<span class="font-mono font-bold text-success">Rp ${(t.debit * 1000000).toLocaleString('id-ID')}</span>` : '-';
        const kreditStr = t.kredit > 0 ? `<span class="font-mono font-bold text-danger">Rp ${(t.kredit * 1000000).toLocaleString('id-ID')}</span>` : '-';
        const balanceStr = `<span class="font-mono font-bold">Rp ${(t.balance * 1000000).toLocaleString('id-ID')}</span>`;
        
        let badgeClass = 'blue';
        if (t.severity === 'danger') badgeClass = 'red';
        else if (t.severity === 'warning') badgeClass = 'orange';
        else if (t.severity === 'normal') badgeClass = 'green';

        return `
          <tr class="${t.severity === 'danger' ? 'bg-danger-light' : ''}">
            <td class="font-mono" style="color: #64748b;">${idx + 1}</td>
            <td class="font-mono" style="white-space: nowrap;">${t.date}</td>
            <td>
              <span class="badge-tag blue" style="font-size: 0.68rem; margin-bottom: 2px;">${t.source}</span><br>
              <span class="font-mono font-bold" style="font-size: 0.76rem;">${t.voucherNo}</span>
            </td>
            <td style="font-size: 0.8rem; line-height: 1.4;">
              <strong>${t.description}</strong>
            </td>
            <td class="text-right">${debitStr}</td>
            <td class="text-right">${kreditStr}</td>
            <td class="text-right">${balanceStr}</td>
            <td><span class="badge-tag ${badgeClass}">${t.flag}</span></td>
            <td class="text-right">
              ${t.severity === 'danger' || t.severity === 'warning' 
                ? `<button class="btn btn-outline-danger btn-xs" onclick="openCaseDetail('AL-2024-001')"><i class="fa-solid fa-folder-open"></i> Dossier</button>`
                : `<button class="btn btn-secondary btn-xs" onclick="showToast('Verifikasi baris #${idx+1} valid.')">Audit</button>`}
            </td>
          </tr>
        `;
      }).join('')}
    </tbody>
  `;
}

// --- CHART.JS INITIALIZATION ---
function initAllCharts() {
  // 1. Off-Hours Journal Line Chart
  const ctxOff = document.getElementById('chart-fin-offhours');
  if (ctxOff) {
    chartInstances['chart-fin-offhours'] = new Chart(ctxOff, {
      type: 'line',
      data: {
        labels: ['00:00', '04:00', '08:00', '12:00', '16:00', '20:00', '23:45'],
        datasets: [
          { label: 'Normal Working Hours Baseline', data: [0, 0, 45, 120, 95, 10, 0], borderColor: '#3B82F6', tension: 0.3 },
          { label: 'Off-Hours Anomalies (Night Posting)', data: [12, 8, 0, 0, 0, 18, 38], borderColor: '#D92525', backgroundColor: 'rgba(217, 37, 37, 0.15)', fill: true, tension: 0.4 }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }

  // 2. Vendor Scatter Plot
  const ctxScatter = document.getElementById('chart-proc-scatter');
  if (ctxScatter) {
    chartInstances['chart-proc-scatter'] = new Chart(ctxScatter, {
      type: 'scatter',
      data: {
        datasets: [
          {
            label: 'Kritis: PT. MAS (Double TF) & PT. Perkasa Kuat (Dirut)',
            data: [{x: 3.84, y: 9.8}, {x: 0.055, y: 9.6}],
            backgroundColor: '#D92525',
            pointRadius: 9
          },
          {
            label: 'Perlu Audit: PT. AADC (Pinjaman 5M) & PT. Ninja',
            data: [{x: 7.41, y: 9.2}, {x: 0.42, y: 7.4}],
            backgroundColor: '#F59E0B',
            pointRadius: 7
          },
          {
            label: 'Wajar: Koperasi IBL, PT. Pahala',
            data: [{x: 0.019, y: 3.0}, {x: 0.059, y: 4.5}],
            backgroundColor: '#10B981',
            pointRadius: 5
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { title: { display: true, text: 'Nilai Transaksi (Rp Miliar)' } },
          y: { title: { display: true, text: 'Tingkat Probabilitas Anomali / ELA Score' } }
        }
      }
    });
  }

  // 7a. Executive Recovery Line
  const ctxExecRec = document.getElementById('chart-exec-recovery');
  if (ctxExecRec) {
    chartInstances['chart-exec-recovery'] = new Chart(ctxExecRec, {
      type: 'line',
      data: {
        labels: ['Jan 04', 'Jan 05 (Double TF)', 'Jan 08', 'Jan 09', 'Jan 10 (Restitusi)', 'Jan 11 (Pinjaman 5M)'],
        datasets: [
          { label: 'Total Kasus Terdeteksi (Rp Miliar)', data: [0.25, 1.23, 1.93, 2.43, 7.33, 12.33], borderColor: '#D92525', backgroundColor: 'rgba(217, 37, 37, 0.15)', fill: true, tension: 0.3 },
          { label: 'Dana Tercegah / Hold Berhasil (Rp Miliar)', data: [0.25, 1.23, 1.23, 1.23, 6.13, 11.13], borderColor: '#10B981', backgroundColor: 'rgba(16, 185, 129, 0.15)', fill: true, tension: 0.3 }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false }
    });
  }

  // 7b. Regional Risk Heatmap Bar
  const ctxExecReg = document.getElementById('chart-exec-regional');
  if (ctxExecReg) {
    chartInstances['chart-exec-regional'] = new Chart(ctxExecReg, {
      type: 'bar',
      data: {
        labels: ['Stockpile ISP Titan Km 107', 'Pit Tambang Kalibata', 'Hauling IPD Alam4 PT. MAS', 'Stockpile Pelangi Ninja', 'Franco Baturaja KDM'],
        datasets: [{
          label: 'Skor Risiko Site Tambang',
          data: [96, 91, 88, 74, 45],
          backgroundColor: ['#D92525', '#D92525', '#D92525', '#EA580C', '#10B981'],
          borderRadius: 6
        }]
      },
      options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
  }
}

// --- NAVIGATION & ROUTING ---
function initNavigation() {
  // 1. Standalone nav buttons (Modul 6-10)
  document.querySelectorAll('.sidebar-nav .nav-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const targetView = btn.getAttribute('data-view');
      if (targetView) {
        document.querySelectorAll('.nav-group').forEach(g => g.classList.remove('open'));
        switchView(targetView);
      }
    });
  });

  // 2. Collapsible nav group headers (Modul 1-5)
  document.querySelectorAll('.nav-group-header').forEach(header => {
    header.addEventListener('click', (e) => {
      e.stopPropagation();
      const parent = header.closest('.nav-group');
      const groupName = header.getAttribute('data-group');
      const isOpen = parent.classList.contains('open');

      document.querySelectorAll('.nav-group').forEach(g => g.classList.remove('open'));

      if (!isOpen) {
        parent.classList.add('open');
      }

      if (groupName) {
        const defaultSub = state.currentSubView[groupName] || (groupName === 'financial' ? 'gl' : groupName === 'procurement' ? 'matrix' : groupName === 'tax' ? 'equalization' : groupName === 'approvals' ? 'sod_matrix' : 'timbangan');
        switchView(groupName, defaultSub);
      }
    });
  });

  // 3. Sub-items inside sidebar accordion
  document.querySelectorAll('.nav-sub-items .sub-item').forEach(sub => {
    sub.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const targetView = sub.getAttribute('data-view');
      const targetSub = sub.getAttribute('data-sub');
      if (targetView && targetSub) switchView(targetView, targetSub);
    });
  });

  // 4. In-module horizontal subnav tab buttons
  document.querySelectorAll('.module-subnav-bar .subnav-tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const targetSub = btn.getAttribute('data-sub');
      const parentModule = btn.closest('.view-panel').id.replace('view-', '');
      if (parentModule && targetSub) switchView(parentModule, targetSub);
    });
  });

  // 5. Sidebar toggle button
  const sidebarToggle = document.getElementById('sidebar-toggle');
  if (sidebarToggle) {
    sidebarToggle.addEventListener('click', (e) => {
      e.stopPropagation();
      if (window.innerWidth <= 1024) {
        document.body.classList.toggle('sidebar-mobile-open');
      } else {
        document.body.classList.toggle('sidebar-collapsed');
      }
      
  // Trigger instant Chart.js resize for mobile viewports
  setTimeout(() => {
    if (window.chartInstances) {
      Object.values(window.chartInstances).forEach(chart => {
        if (chart && typeof chart.resize === 'function') {
          chart.resize();
        }
      });
    }
    window.dispatchEvent(new Event('resize'));
  }, 60);

    });
  }

  // 6. Sidebar backdrop click
  const backdrop = document.getElementById('sidebar-backdrop');
  if (backdrop) {
    backdrop.addEventListener('click', () => {
      document.body.classList.remove('sidebar-mobile-open');
    });
  }
}

function switchView(viewName, subViewName) {
  if (window.innerWidth <= 1024) {
    document.body.classList.remove('sidebar-mobile-open');
  }

  state.currentView = viewName;
  const defaultSubMap = {
    financial: 'gl',
    procurement: 'matrix',
    tax: 'equalization',
    approvals: 'sod_matrix',
    operational: 'timbangan'
  };
  const sub = subViewName || state.currentSubView[viewName] || defaultSubMap[viewName] || 'gl';
  state.currentSubView[viewName] = sub;

  // 1. Switch View Panel
  document.querySelectorAll('.view-panel').forEach(p => p.classList.remove('active'));
  const targetPanel = document.getElementById(`view-${viewName}`);
  if (targetPanel) targetPanel.classList.add('active');

  // 2. Switch Sub-View Panes
  const prefixMap = {
    financial: 'fin-sub-',
    procurement: 'proc-sub-',
    tax: 'tax-sub-',
    approvals: 'app-sub-',
    operational: 'ops-sub-'
  };

  const prefix = prefixMap[viewName];
  if (prefix && targetPanel) {
    targetPanel.querySelectorAll('.sub-view-pane').forEach(pane => pane.classList.remove('active'));
    let targetPane = document.getElementById(`${prefix}${sub}`);
    if (!targetPane) {
      targetPane = targetPanel.querySelector('.sub-view-pane');
    }
    if (targetPane) targetPane.classList.add('active');

    // Update in-module subnav tab buttons
    const subnavBar = document.getElementById(`subnav-${viewName}`);
    if (subnavBar) {
      subnavBar.querySelectorAll('.subnav-tab-btn').forEach(b => {
        b.classList.remove('active');
        if (b.getAttribute('data-sub') === sub) b.classList.add('active');
      });
      const activeTab = subnavBar.querySelector('.subnav-tab-btn.active');
      if (activeTab) {
        activeTab.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
      }
    }

    // Update sidebar sub-item active state
    document.querySelectorAll('.nav-sub-items .sub-item').forEach(si => si.classList.remove('active'));
    const currentSubItem = document.querySelector(`.nav-sub-items .sub-item[data-view="${viewName}"][data-sub="${sub}"]`);
    if (currentSubItem) currentSubItem.classList.add('active');

    // Make sure parent group is open
    const parentGroup = document.querySelector(`button[data-group="${viewName}"]`);
    if (parentGroup) {
      document.querySelectorAll('.nav-group').forEach(g => g.classList.remove('open'));
      parentGroup.closest('.nav-group').classList.add('open');
    }
  }

  // 3. Update main sidebar active state
  document.querySelectorAll('.sidebar-nav .nav-btn').forEach(b => {
    b.classList.remove('active');
    if (b.getAttribute('data-view') === viewName) b.classList.add('active');
  });

  // 4. Update Breadcrumb
  const breadcrumb = document.getElementById('current-breadcrumb');
  if (breadcrumb) {
    breadcrumb.textContent = getModuleTitle(viewName, sub);
  }

  
  // Trigger instant Chart.js resize for mobile viewports
  setTimeout(() => {
    if (window.chartInstances) {
      Object.values(window.chartInstances).forEach(chart => {
        if (chart && typeof chart.resize === 'function') {
          chart.resize();
        }
      });
    }
    window.dispatchEvent(new Event('resize'));
  }, 60);

  renderModuleCharts(viewName, sub); window.scrollTo({ top: 0, behavior: 'smooth' });
}

function getModuleTitle(view, sub) {
  const titles = {
    financial: '1. Buku Kas & Bank Mandiri Forensics',
    procurement: '2. Kontraktor & Vendor Hauling Forensics',
    tax: '3. Pajak & Ekualisasi DJP 3-Way (JendelaTax)',
    approvals: '4. Kontrol Internal, SoD & Override Log',
    operational: '5. Timbangan & Logistik Stockpile Tambang',
    case_mgmt: '6. Antrean Kasus Triase Investigasi',
    executive: '7. Executive Dashboard POJK 12/2024',
    investigation_detail: '8. Digital Dossier Workspace (AL-001)',
    evidence_mgmt: '9. Digital Evidence Vault (SHA-256)',
    user_access: '10. Tata Kelola User Access & RBAC'
  };
  return titles[view] || 'Forensic Dashboard';
}

// --- MODALS & ACTIONS ---
function openCaseDetail(caseId) {
  switchView('investigation_detail', 'overview');
  showToast(`Membuka berkas perkara digital dossier: ${caseId}`);
}

function openOcrInspectorModal(invoiceId) {
  const m = document.getElementById('modal-ocr-inspector');
  if (m) m.classList.add('active');
}

function closeModal(modalId) {
  const m = document.getElementById(modalId);
  if (m) m.classList.remove('active');
}

function openNewCaseModal() {
  const title = prompt('Masukkan judul berkas perkara investigasi baru:', 'Anomali Jasa Hauling Stockpile Batubara');
  if (title) {
    state.cases.unshift({
      id: `AL-2024-${Date.now().toString().slice(-4)}`,
      title: title,
      module: 'Procurement & Hauling',
      entity: 'PT. MAS (Mitra Alam Sentosa)',
      amount: 'Rp 980.000.000',
      severity: 'High',
      status: 'New',
      investigator: 'Budi Santoso, CFE'
    });
    showToast('Berkas perkara baru berhasil didaftarkan ke antrean triase.');
  }
}

function openUploadEvidenceModal() {
  showToast('Pilih berkas bukti digital untuk diverifikasi integritas SHA-256...');
}

function showToast(message) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerHTML = `<i class="fa-solid fa-shield-halved text-danger"></i> <span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(100%)';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}


// --- DYNAMIC LAZY CHART RENDERER (GUARANTEED 100% MOBILE FIT) ---
window.chartStore = {};

function renderModuleCharts(viewName, subName) {
  setTimeout(() => {
    // 1. Module Financial Charts
    if (viewName === 'financial') {
      const ctxOff = document.getElementById('chart-fin-offhours');
      if (ctxOff) {
        if (window.chartStore['chart-fin-offhours']) window.chartStore['chart-fin-offhours'].destroy();
        window.chartStore['chart-fin-offhours'] = new Chart(ctxOff, {
          type: 'line',
          data: {
            labels: ['00:00', '04:00', '08:00', '12:00', '16:00', '20:00', '23:45'],
            datasets: [
              { label: 'Normal Baseline', data: [0, 0, 45, 120, 95, 10, 0], borderColor: '#3B82F6', tension: 0.3 },
              { label: 'Off-Hours Anomalies', data: [12, 8, 0, 0, 0, 18, 38], borderColor: '#D92525', backgroundColor: 'rgba(217, 37, 37, 0.15)', fill: true, tension: 0.4 }
            ]
          },
          options: { responsive: true, maintainAspectRatio: false, scales: { x: { ticks: { font: { size: 9 } } }, y: { ticks: { font: { size: 9 } } } } }
        });
      }
    }

    // 2. Module Procurement Charts
    if (viewName === 'procurement') {
      const ctxScatter = document.getElementById('chart-proc-scatter');
      if (ctxScatter) {
        if (window.chartStore['chart-proc-scatter']) window.chartStore['chart-proc-scatter'].destroy();
        window.chartStore['chart-proc-scatter'] = new Chart(ctxScatter, {
          type: 'scatter',
          data: {
            datasets: [
              { label: 'Kritis', data: [{x: 8.8, y: 9.2}, {x: 8.2, y: 8.7}], backgroundColor: '#D92525', pointRadius: 7 },
              { label: 'Wajar', data: [{x: 1.5, y: 2.0}, {x: 2.1, y: 1.8}], backgroundColor: '#10B981', pointRadius: 5 }
            ]
          },
          options: { responsive: true, maintainAspectRatio: false, scales: { x: { ticks: { font: { size: 9 } } }, y: { ticks: { font: { size: 9 } } } } }
        });
      }
    }

    // 3. Module Operational (POS Hourly Chart)
    if (viewName === 'operational') {
      const ctxPos = document.getElementById('chart-ops-pos');
      if (ctxPos) {
        if (window.chartStore['chart-ops-pos']) window.chartStore['chart-ops-pos'].destroy();
        window.chartStore['chart-ops-pos'] = new Chart(ctxPos, {
          type: 'line',
          data: {
            labels: ['08:00', '10:00', '12:00', '14:00', '16:00', '18:00', '21:00'],
            datasets: [{
              label: 'Spike Anomali Void Pasca-Customer',
              data: [2, 3, 5, 4, 8, 14, 45],
              borderColor: '#D92525',
              backgroundColor: 'rgba(217, 37, 37, 0.15)',
              fill: true,
              tension: 0.4
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: {
                labels: { boxWidth: 10, font: { size: 10 } }
              }
            },
            scales: {
              x: { ticks: { font: { size: 9 }, maxRotation: 0, autoSkip: true } },
              y: { ticks: { font: { size: 9 } } }
            }
          }
        });
      }
    }

    // 4. Module Executive Charts
    if (viewName === 'executive') {
      const ctxExecRec = document.getElementById('chart-exec-recovery');
      if (ctxExecRec) {
        if (window.chartStore['chart-exec-recovery']) window.chartStore['chart-exec-recovery'].destroy();
        window.chartStore['chart-exec-recovery'] = new Chart(ctxExecRec, {
          type: 'line',
          data: {
            labels: ['Kasus 1', 'Kasus 2', 'Kasus 3', 'Kasus 4', 'Kasus 5'],
            datasets: [
              { label: 'Terdeteksi (Miliar)', data: [1.2, 2.5, 4.8, 8.1, 12.3], borderColor: '#D92525', tension: 0.3 },
              { label: 'Tercegah (Miliar)', data: [1.2, 2.5, 4.0, 7.5, 11.1], borderColor: '#10B981', tension: 0.3 }
            ]
          },
          options: { responsive: true, maintainAspectRatio: false, scales: { x: { ticks: { font: { size: 9 } } }, y: { ticks: { font: { size: 9 } } } } }
        });
      }
    }
  }, 40);
}
