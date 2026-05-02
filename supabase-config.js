// Supabase configuration for Bestie Visit.
// Isi url dan publishableKey dari Supabase Dashboard > Project Settings > API Keys.
// Jangan pernah memasukkan secret key/service_role key di file frontend ini.
window.RB_SUPABASE_CONFIG = {
  enabled: false,
  url: 'https://wsyttlpvwumgdhyzvvbq.supabase.co',
  publishableKey: 'sb_publishable_mDEPMhW-YUQImNe0xauF2w_Plw4QP-E',

  // Browser bundle; app ini tidak memakai build step/npm.
  bundleUrl: 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2',

  tables: {
    appSettings: 'app_settings',
    monitor: 'monitor_visits',
    manualRequests: 'manual_store_requests',
    presence: 'monitor_presence'
  },

  // HOME info/welcome + monitor admin akan refresh berkala.
  pollMs: 5000,
  monitorLimit: 500
};
