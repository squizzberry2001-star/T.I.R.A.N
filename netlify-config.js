// Netlify configuration for Bestie Visit.
// Database sync memakai Netlify Functions + Netlify Blobs.
// Deploy file netlify/functions/rbv-data.mjs dan package.json agar function aktif.

window.RB_NETLIFY_CONFIG = {
  enabled: true,
  // Default path jika web sudah di-host di Netlify yang sama.
  functionPath: 'https://adminpanelbestie.netlify.app/.netlify/functions/rbv-data?action=listAppSettings',
  // Isi hanya jika function dipasang di domain Netlify berbeda. Contoh: 'https://nama-site.netlify.app'
  baseUrl: 'https://adminpanelbestie.netlify.app/',
  // Optional. Kalau kamu set environment variable RBV_ADMIN_TOKEN di Netlify, isi token yang sama di sini.
  // Untuk mode simple tanpa login, biarkan kosong dan jangan set RBV_ADMIN_TOKEN.
  adminToken: '',
  pollMs: 5000,
  monitorLimit: 500
};
