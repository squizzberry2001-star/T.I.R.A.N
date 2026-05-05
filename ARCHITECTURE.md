# Arsitektur v3

## Static Frontend

Semua fitur utama berjalan di browser:

- ZIP extraction
- static preview
- GitHub auth token
- push ZIP
- rollback
- repo terminal
- repo editor

## Repo Terminal

Repo Terminal bukan shell OS. Ia adalah command layer di atas GitHub Git Database API.

Command:

- `pull`: reload tree dari branch main
- `ls`: render tree
- `cd`: ubah cwd client-side
- `open`: fetch blob dari GitHub dan isi editor
- `save`: stage isi editor client-side
- `mkdir`: stage `.gitkeep`
- `touch`: stage file kosong
- `rm`: stage delete
- `commit`: membuat blob/tree/commit baru dan update `refs/heads/main`

## Alur commit terminal

1. Ambil current ref `main`.
2. Ambil commit dan base tree.
3. Untuk staged save/add, buat blob baru.
4. Untuk staged delete, masukkan tree entry dengan `sha: null`.
5. Buat tree baru dengan `base_tree`.
6. Buat commit baru.
7. PATCH `refs/heads/main`.
8. Simpan rollback point ke localStorage.

## Mobile layout

- CSS memakai media query `max-width` dan `orientation: landscape`.
- Tombol command mengurangi kebutuhan mengetik simbol dan command panjang.
- Area terminal dan editor memakai tinggi viewport adaptif.
