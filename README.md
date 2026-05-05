# ZIP to GitHub Auto Push v3

Web app pribadi untuk upload ZIP, preview, push ke GitHub branch `main`, rollback, dan edit repository lewat **Repo Terminal + Editor** yang nyaman dipakai di HP.

## Fitur utama

- Login GitHub via token.
- Upload dan extract ZIP di browser.
- Preview struktur file dan source code.
- Preview static website dari `index.html` memakai Service Worker + Cache API.
- Push ZIP ke branch `main`.
- File path sama ditimpa; file lain di repo tetap ada.
- Rollback last push atau rollback ke commit terpilih.
- Repo Terminal + Editor berdasarkan repo yang dipilih.

## Fitur Repo Terminal

Setelah login dan pilih repository:

1. Klik **Load Workspace Repo**.
2. Gunakan tombol cepat atau command:
   - `help`
   - `pwd`
   - `ls`
   - `cd folder`
   - `open file`
   - `new file`
   - `save`
   - `touch file`
   - `mkdir folder`
   - `rm file`
   - `status`
   - `discard`
   - `commit pesan commit`
   - `pull`
   - `clear`
3. Edit isi file di editor.
4. Klik **Stage Save**.
5. Klik **Commit Staged ke main**.

## Kenapa bukan terminal Linux asli?

GitHub repository tidak menyediakan shell terminal langsung. Terminal di app ini adalah terminal kontrol repository berbasis GitHub API. Ini bisa berjalan di Vercel static app tanpa backend.

Untuk command runtime sungguhan seperti `npm install`, `php artisan`, atau `python app.py`, gunakan runner Docker opsional di folder `runner/`.

## Deploy ke Vercel

1. Push folder ini ke GitHub.
2. Import ke Vercel.
3. Framework preset: `Other`.
4. Build command: kosong.
5. Output directory: root/kosong.

## Permission token

Minimal:

- Contents: Read and write
- Metadata: Read-only

Jika mengedit `.github/workflows/*`, token mungkin butuh permission Workflows.

## Mobile UX

- Portrait: terminal dan editor bertumpuk.
- Landscape: terminal dan editor berdampingan.
- Tombol command besar.
- Input terminal sticky.
- Font-size input 16px untuk mengurangi auto-zoom di HP.
