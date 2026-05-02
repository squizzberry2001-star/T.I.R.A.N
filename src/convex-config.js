// Convex realtime/backend configuration for Bestie Visit.
// Development URL dari Convex dashboard/terminal: third-monitor-660.
// Untuk production, ganti deploymentUrl dengan URL production kamu, lalu jalankan npx convex deploy.
window.RB_CONVEX_CONFIG = {
  enabled: true,

  // Realtime WebSocket client URL.
  deploymentUrl: 'https://third-monitor-660.convex.cloud',

  // Optional legacy HTTP fallback URL.
  httpUrl: '',

  // Static browser bundle; app ini tidak memakai build step.
  bundleUrl: 'https://unpkg.com/convex@latest/dist/browser.bundle.js',

  // Secret admin monitor functions.
  monitorQuery: 'monitor:listVisits',
  manualRequestsQuery: 'monitor:listManualStoreRequests',
  presenceQuery: 'monitor:listPresence',
  upsertMutation: 'monitor:upsertVisit',
  upsertManualRequestMutation: 'monitor:upsertManualStoreRequest',
  presenceUpsertMutation: 'monitor:upsertPresence',

  // Linked device + visit sync functions.
  registerDeviceMutation: 'linkedDevices:registerDevice',
  createLinkMutation: 'linkedDevices:createLinkCode',
  acceptLinkMutation: 'linkedDevices:acceptLinkCode',
  linkedDevicesQuery: 'linkedDevices:getLinkedDevices',
  visitUpsertMutation: 'visits:upsertVisit',
  visitDeleteMutation: 'visits:deleteVisit',
  visitListQuery: 'visits:listVisits',

  // Legacy HTTP action paths, only used as fallback if deploymentUrl is empty/unavailable.
  upsertPath: 'monitor/upsertVisit',
  listPath: 'monitor/listVisits',
  upsertManualRequestPath: 'monitor/upsertManualStoreRequest',
  listManualRequestsPath: 'monitor/listManualStoreRequests',
  upsertPresencePath: 'monitor/upsertPresence',
  listPresencePath: 'monitor/listPresence',

  pollMs: 5000,
  token: ''
};
