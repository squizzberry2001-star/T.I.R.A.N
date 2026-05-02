import { getStore } from '@netlify/blobs';

const STORE_NAME = env('RBV_STORE_NAME') || 'bestie-visit-data';
const KEY_MONITOR = 'monitor_visits.json';
const KEY_MANUAL = 'manual_store_requests.json';
const KEY_PRESENCE = 'monitor_presence.json';
const KEY_SETTINGS = 'app_settings.json';
const MAX_MONITOR = Number(env('RBV_MONITOR_LIMIT') || 500);
const MAX_MANUAL = Number(env('RBV_MANUAL_LIMIT') || 500);
const MAX_PRESENCE = Number(env('RBV_PRESENCE_LIMIT') || 300);

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type,Authorization,X-Admin-Token',
  'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
  'Content-Type': 'application/json; charset=utf-8'
};

function env(name) {
  try {
    if (globalThis.Netlify?.env?.get) return globalThis.Netlify.env.get(name) || '';
  } catch (error) {}
  return process.env[name] || '';
}

function json(payload, status = 200) {
  return new Response(JSON.stringify(payload), { status, headers: CORS });
}

function error(message, status = 500) {
  return json({ ok: false, error: message || 'Server error' }, status);
}

function parseLimit(value, fallback) {
  const n = Number(value || fallback);
  return Number.isFinite(n) ? Math.max(1, Math.min(1000, Math.round(n))) : fallback;
}

async function readJSON(store, key, fallback) {
  const value = await store.get(key, { type: 'json', consistency: 'strong' });
  return value == null ? fallback : value;
}

async function writeJSON(store, key, value) {
  await store.setJSON(key, value, { metadata: { updatedAt: new Date().toISOString() } });
}

function cleanText(value, fallback = '') {
  const text = value == null ? '' : String(value).trim();
  return text || fallback;
}

function requestToken(request) {
  const direct = request.headers.get('x-admin-token') || '';
  const auth = request.headers.get('authorization') || '';
  return cleanText(direct || auth.replace(/^Bearer\s+/i, ''));
}

function assertWriteAllowed(request) {
  const expected = cleanText(env('RBV_ADMIN_TOKEN') || env('NETLIFY_ADMIN_TOKEN'));
  if (!expected) return;
  if (requestToken(request) !== expected) {
    const err = new Error('Token admin Netlify salah atau belum dikirim.');
    err.status = 401;
    throw err;
  }
}

function normalizeArray(value) {
  return Array.isArray(value) ? value : [];
}

function sortByUpdatedDesc(rows) {
  return normalizeArray(rows).sort((a, b) => {
    const left = Date.parse(b.updated_at || b.updatedAt || b.last_visit_at || b.created_at || 0) || Number(b.updated_at || b.updatedAt || 0) || 0;
    const right = Date.parse(a.updated_at || a.updatedAt || a.last_visit_at || a.created_at || 0) || Number(a.updated_at || a.updatedAt || 0) || 0;
    return left - right;
  });
}

function mergeByKey(rows, keyName, payload) {
  const key = cleanText(payload?.[keyName]);
  if (!key) return normalizeArray(rows);
  const now = new Date().toISOString();
  const nextPayload = { ...payload, updated_at: payload.updated_at || now };
  const list = normalizeArray(rows).filter((item) => cleanText(item?.[keyName]) !== key);
  list.unshift(nextPayload);
  return list;
}

async function listMonitorVisits(store, requestUrl) {
  const limit = parseLimit(requestUrl.searchParams.get('limit'), MAX_MONITOR);
  const rows = sortByUpdatedDesc(await readJSON(store, KEY_MONITOR, []));
  return json({ ok: true, rows: rows.slice(0, limit) });
}

async function upsertMonitorVisit(store, request) {
  assertWriteAllowed(request);
  const payload = await request.json();
  const visitKey = cleanText(payload.visit_key);
  if (!visitKey) return error('visit_key wajib diisi.', 400);
  const rows = await readJSON(store, KEY_MONITOR, []);
  const nextRows = sortByUpdatedDesc(mergeByKey(rows, 'visit_key', payload)).slice(0, MAX_MONITOR);
  await writeJSON(store, KEY_MONITOR, nextRows);
  return json({ ok: true, row: nextRows.find((item) => cleanText(item.visit_key) === visitKey) || payload });
}

async function listPresence(store, requestUrl) {
  const limit = parseLimit(requestUrl.searchParams.get('limit'), MAX_PRESENCE);
  const rows = sortByUpdatedDesc(await readJSON(store, KEY_PRESENCE, []));
  return json({ ok: true, rows: rows.slice(0, limit) });
}

async function upsertPresence(store, request) {
  assertWriteAllowed(request);
  const payload = await request.json();
  const sessionId = cleanText(payload.session_id);
  if (!sessionId) return error('session_id wajib diisi.', 400);
  const now = new Date().toISOString();
  const rows = await readJSON(store, KEY_PRESENCE, []);
  const nextPayload = { ...payload, session_id: sessionId, last_seen_at: payload.last_seen_at || now, updated_at: payload.updated_at || now };
  const nextRows = sortByUpdatedDesc(mergeByKey(rows, 'session_id', nextPayload)).slice(0, MAX_PRESENCE);
  await writeJSON(store, KEY_PRESENCE, nextRows);
  return json({ ok: true, row: nextRows.find((item) => cleanText(item.session_id) === sessionId) || nextPayload });
}

async function listManualRequests(store) {
  const rows = sortByUpdatedDesc(await readJSON(store, KEY_MANUAL, []));
  return json({ ok: true, rows: rows.slice(0, MAX_MANUAL) });
}

async function upsertManualRequest(store, request) {
  assertWriteAllowed(request);
  const payload = await request.json();
  const requestId = cleanText(payload.request_id);
  if (!requestId) return error('request_id wajib diisi.', 400);
  const rows = await readJSON(store, KEY_MANUAL, []);
  const nextRows = sortByUpdatedDesc(mergeByKey(rows, 'request_id', payload)).slice(0, MAX_MANUAL);
  await writeJSON(store, KEY_MANUAL, nextRows);
  return json({ ok: true, row: nextRows.find((item) => cleanText(item.request_id) === requestId) || payload });
}

async function listAppSettings(store, requestUrl) {
  const settings = await readJSON(store, KEY_SETTINGS, {});
  const keys = cleanText(requestUrl.searchParams.get('keys')).split(',').map((item) => item.trim()).filter(Boolean);
  const rows = Object.entries(settings)
    .filter(([key]) => !keys.length || keys.includes(key))
    .map(([key, value]) => ({ key, config_key: key, ...(value && typeof value === 'object' ? value : { payload: value }) }));
  return json({ ok: true, rows });
}

async function setAppSetting(store, request) {
  assertWriteAllowed(request);
  const body = await request.json();
  const key = cleanText(body.key || body.config_key);
  if (!key) return error('key wajib diisi.', 400);
  const settings = await readJSON(store, KEY_SETTINGS, {});
  settings[key] = {
    key,
    config_key: key,
    payload: body.payload || {},
    updatedAt: new Date().toISOString(),
    updated_at: new Date().toISOString(),
    updatedBy: cleanText(body.updatedBy || body.updated_by || 'web'),
    updated_by: cleanText(body.updatedBy || body.updated_by || 'web')
  };
  await writeJSON(store, KEY_SETTINGS, settings);
  return json({ ok: true, row: settings[key] });
}

export default async function handler(request) {
  if (request.method === 'OPTIONS') return new Response('', { status: 204, headers: CORS });
  const requestUrl = new URL(request.url);
  const action = requestUrl.searchParams.get('action') || '';
  const store = getStore(STORE_NAME);
  try {
    if (request.method === 'GET' && action === 'listMonitorVisits') return listMonitorVisits(store, requestUrl);
    if (request.method === 'POST' && action === 'upsertMonitorVisit') return upsertMonitorVisit(store, request);
    if (request.method === 'GET' && action === 'listPresence') return listPresence(store, requestUrl);
    if (request.method === 'POST' && action === 'upsertPresence') return upsertPresence(store, request);
    if (request.method === 'GET' && action === 'listManualRequests') return listManualRequests(store);
    if (request.method === 'POST' && action === 'upsertManualRequest') return upsertManualRequest(store, request);
    if (request.method === 'GET' && action === 'listAppSettings') return listAppSettings(store, requestUrl);
    if (request.method === 'POST' && action === 'setAppSetting') return setAppSetting(store, request);
    return error('Action tidak dikenal.', 404);
  } catch (err) {
    console.error('RBV Netlify function error:', err);
    return error(err?.message || 'Netlify function gagal.', err?.status || 500);
  }
}
