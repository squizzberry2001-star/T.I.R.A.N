# T.I.R.A.N

T.I.R.A.N adalah UI/UX web AI bergaya chat modern yang siap dipakai untuk GitHub + Vercel.

## Yang sudah ada

- Tampilan chat AI modern ala workspace assistant
- Multi-chat dengan riwayat percakapan tersimpan di browser (`localStorage`)
- Sidebar riwayat + pencarian percakapan
- Streaming response real-time dari server route
- Model selector cepat di composer
- Pengaturan model, tema, dan system prompt
- Regenerate jawaban terakhir
- Copy per-message dan copy per-code-block
- Export chat ke Markdown
- Responsive untuk desktop dan mobile
- API key aman di server via environment variable

## Stack

- Next.js App Router
- TypeScript
- CSS custom tanpa dependensi UI tambahan
- OpenAI Chat Completions API via server route

## Jalankan lokal

```bash
npm install
npm run dev
```

Buat file `.env.local`:

```env
OPENAI_API_KEY=your_openai_api_key
OPENAI_MODEL=gpt-5-chat-latest
```

## Deploy ke GitHub dan Vercel

1. Buat repository baru di GitHub.
2. Upload semua isi folder ini ke repository.
3. Login ke Vercel lalu pilih **Add New Project**.
4. Import repository GitHub Anda.
5. Tambahkan environment variable:
   - `OPENAI_API_KEY`
   - `OPENAI_MODEL` (opsional)
6. Deploy.

## Struktur penting

- `app/page.tsx` → halaman utama
- `components/tiran-app.tsx` → logika UI utama
- `components/message-content.tsx` → render konten dan code block
- `app/api/chat/route.ts` → proxy aman ke OpenAI
- `app/globals.css` → seluruh styling

## Catatan

Versi ini fokus pada pengalaman chat inti seperti GPT-style assistant. Kalau Anda ingin, tahap berikutnya paling masuk akal adalah menambahkan:

- login/auth
- database untuk sinkron riwayat lintas perangkat
- upload file/gambar
- voice input/output
- prompt templates
- admin analytics
- billing / subscription

