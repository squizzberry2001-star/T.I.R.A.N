# Regional Bestie Visit Report - React Revamp 2026

Aplikasi audit Regional Bestie berbasis React + Tailwind CSS.

## Isi utama
- `index.html` - entry point aplikasi.
- `src/app.jsx` - source React utama.
- `src/app.js` - hasil compile dari `app.jsx` untuk browser.
- `src/theme.css` dan `src/tailwind.generated.css` - styling UI.
- `src/pdf-generator.js` dan `src/pdf-template-assets.js` - generator PDF sesuai template.
- `data.js` dan `store-master-data.js` - data bestie dan master store.
- `ca-assignment-export.js` dan `jszip.min.js` - export Excel CA Assignment.
- `convex-config.js` - konfigurasi monitor admin Convex.
- `convex-monitor-example.js` - contoh HTTP actions Convex.

## Catatan PDF
Generator PDF memakai ukuran landscape 16:9 seperti template. Slide QSC/Famitrack menampilkan 2 foto besar berdampingan. Tabel OPI/QSC memakai font 12.5, margin konsisten, dan otomatis membuat halaman baru saat konten penuh.

## Mengaktifkan monitor admin Convex
1. Deploy HTTP actions di Convex sesuai contoh `convex-monitor-example.js`.
2. Isi `convex-config.js` dengan `enabled: true`, `httpUrl`, dan path endpoint.
3. Panel admin dibuka dari logo/title Bestie Visit dengan klik 10x, lalu PIN `607090`.

## Revamp v5
- Mobile navbar dibuat menyamping di bawah dengan padding aman agar tidak menutup konten.
- Toggle section hanya berfungsi hide/unhide slide; konten QSC, Observation, dan Evidence tampil hanya saat toggle aktif.
- Crop photo otomatis muncul setelah upload, background halaman dikunci, zoom memakai pinch/gesture tanpa slider.
- Rich text dipindah ke bawah textbox dengan state hover/active dan perbaikan command bullet/number.
- PDF menambahkan footer watermark `GENERATE BY BESTIE VISIT WEB REPORT` di semua halaman, General Information memakai label Visitor, dan grid evidence maksimal 6 foto per halaman.

## Revamp v6
- Mobile header logo diganti menjadi quick section sticky yang ringkas.
- Semua label `Section 1 - Section 6`, subtitle panduan, dan narasi form yang tidak perlu dihapus dari tampilan utama.
- Dashboard dibuat lebih minimalis, dilengkapi tombol info blinking untuk panduan Add to Home Screen/PWA.
- Dropdown Nama Bestie dan Store di modal kunjungan baru serta Section 1 memakai native select agar lebih stabil di mobile.
- Request toko manual tersimpan sebagai approval queue di Panel Admin.
- Store Leader tidak lagi auto-fill dari master data; input manual.
- Crop/marker photo popup diperbaiki agar proporsional dan body tidak ikut scroll.
- Preview PDF diganti menjadi canvas-based preview menggunakan PDF.js agar ukuran slide mengikuti layar mobile.

## Revamp v7
- Memperbaiki alignment dashboard home, field select, dan date input di mobile.
- Rich text toolbar dibuat lebih proper; bullet dan number sekarang muncul di editor saat tombol dipilih.
- Popup crop & marker dibuat proporsional dengan kanvas 1:1 dan gesture tetap smooth.
- Preview PDF diperbaiki agar tidak loading/blinking berulang di mobile.
- PDF evidence menggunakan foto rasio 1:1 yang lebih jelas pada grid.
