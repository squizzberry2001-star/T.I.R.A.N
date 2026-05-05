# v3 Changes

## Ditambahkan

- Repo Terminal + Editor.
- Load workspace repository terpilih.
- Command `help`, `pwd`, `ls`, `cd`, `open`, `new`, `save`, `touch`, `mkdir`, `rm`, `status`, `discard`, `commit`, `pull`, `clear`.
- Dropdown file repo.
- Tombol cepat untuk HP.
- Layout portrait/horizontal.
- Commit staged dari editor ke main.

## Batasan

- Ini bukan shell Linux asli.
- Tidak bisa menjalankan `npm install`, `php artisan`, `python app.py` langsung di repo GitHub.
- Untuk command runtime sungguhan tetap perlu runner Docker.
- Binary file tidak bisa diedit di textarea.
- GitHub tidak menyimpan folder kosong, jadi `mkdir` membuat `.gitkeep`.
