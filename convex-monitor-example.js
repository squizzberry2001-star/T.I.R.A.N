/*
Convex realtime backend example for Bestie Visit Secret Panel
=============================================================

Use this when you want the admin monitor to update realtime across devices.
Convex realtime requires normal Convex queries/mutations. The web app subscribes to:
- monitor:listVisits
- monitor:listManualStoreRequests
- monitor:listPresence

And writes through mutations:
- monitor:upsertVisit
- monitor:upsertManualStoreRequest
- monitor:upsertPresence

Recommended setup:
1. In your project terminal run:
   npm install convex
   npx convex dev
2. Create these files in the generated convex/ folder.
3. Deploy with:
   npx convex deploy
4. Copy your production URL from Convex dashboard, usually https://xxxxx.convex.cloud.
5. Put it in convex-config.js as deploymentUrl and set enabled: true.

File: convex/schema.ts
----------------------
import { defineSchema, defineTable } from "convex/server";
import { v } from "convex/values";

export default defineSchema({
  regional_bestie_visits: defineTable({
    visit_key: v.string(),
    bestie_name: v.string(),
    store_name: v.string(),
    store_code: v.optional(v.string()),
    visit_date: v.string(),
    total_visits: v.number(),
    last_visit_at: v.string(),
    updated_at: v.string(),
    session_id: v.optional(v.string()),
    event_type: v.optional(v.string()),
    page_url: v.optional(v.string()),
    user_agent: v.optional(v.string())
  })
    .index("by_visit_key", ["visit_key"])
    .index("by_updated_at", ["updated_at"]),

  monitor_presence: defineTable({
    session_id: v.string(),
    bestie_name: v.string(),
    store_name: v.string(),
    store_code: v.optional(v.string()),
    visit_id: v.optional(v.string()),
    active_screen: v.optional(v.string()),
    is_online: v.optional(v.boolean()),
    last_seen_at: v.string(),
    updated_at: v.string(),
    page_url: v.optional(v.string()),
    user_agent: v.optional(v.string())
  })
    .index("by_session_id", ["session_id"])
    .index("by_updated_at", ["updated_at"]),

  manual_store_requests: defineTable({
    request_id: v.string(),
    status: v.string(),
    created_at: v.number(),
    updated_at: v.number(),
    bestie_name: v.optional(v.string()),
    store_name: v.string(),
    store_code: v.optional(v.string()),
    address: v.optional(v.string()),
    note: v.optional(v.string()),
    session_id: v.optional(v.string()),
    page_url: v.optional(v.string()),
    user_agent: v.optional(v.string())
  })
    .index("by_request_id", ["request_id"])
    .index("by_status", ["status"])
    .index("by_updated_at", ["updated_at"])
});

File: convex/monitor.ts
-----------------------
import { mutation, query } from "./_generated/server";
import { v } from "convex/values";

const text = (value: unknown, fallback = "") => {
  if (value === undefined || value === null) return fallback;
  const clean = String(value).trim();
  return clean || fallback;
};

const num = (value: unknown, fallback = Date.now()) => {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
};

export const listVisits = query({
  args: {},
  handler: async (ctx) => {
    return await ctx.db
      .query("regional_bestie_visits")
      .withIndex("by_updated_at")
      .order("desc")
      .take(500);
  }
});

export const upsertVisit = mutation({
  args: { payload: v.any() },
  handler: async (ctx, { payload }) => {
    const visit_key = text(payload.visit_key);
    if (!visit_key) throw new Error("visit_key wajib diisi");

    const now = new Date().toISOString();
    const doc = {
      visit_key,
      bestie_name: text(payload.bestie_name, "-"),
      store_name: text(payload.store_name, "-"),
      store_code: text(payload.store_code),
      visit_date: text(payload.visit_date, now.slice(0, 10)),
      total_visits: Number(payload.total_visits || 1),
      last_visit_at: text(payload.last_visit_at, now),
      updated_at: text(payload.updated_at, now),
      session_id: text(payload.session_id),
      event_type: text(payload.event_type),
      page_url: text(payload.page_url),
      user_agent: text(payload.user_agent)
    };

    const existing = await ctx.db
      .query("regional_bestie_visits")
      .withIndex("by_visit_key", (q) => q.eq("visit_key", visit_key))
      .unique();

    if (existing) {
      await ctx.db.patch(existing._id, doc);
      return { ok: true, id: existing._id, updated: true };
    }

    const id = await ctx.db.insert("regional_bestie_visits", doc);
    return { ok: true, id, inserted: true };
  }
});

export const listPresence = query({
  args: {},
  handler: async (ctx) => {
    return await ctx.db
      .query("monitor_presence")
      .withIndex("by_updated_at")
      .order("desc")
      .take(300);
  }
});

export const upsertPresence = mutation({
  args: { payload: v.any() },
  handler: async (ctx, { payload }) => {
    const session_id = text(payload.session_id);
    if (!session_id) throw new Error("session_id wajib diisi");

    const now = new Date().toISOString();
    const doc = {
      session_id,
      bestie_name: text(payload.bestie_name, "Belum pilih bestie"),
      store_name: text(payload.store_name, "Belum pilih store"),
      store_code: text(payload.store_code),
      visit_id: text(payload.visit_id),
      active_screen: text(payload.active_screen, "home"),
      is_online: payload.is_online !== false,
      last_seen_at: text(payload.last_seen_at, now),
      updated_at: text(payload.updated_at, now),
      page_url: text(payload.page_url),
      user_agent: text(payload.user_agent)
    };

    const existing = await ctx.db
      .query("monitor_presence")
      .withIndex("by_session_id", (q) => q.eq("session_id", session_id))
      .unique();

    if (existing) {
      await ctx.db.patch(existing._id, doc);
      return { ok: true, id: existing._id, updated: true };
    }

    const id = await ctx.db.insert("monitor_presence", doc);
    return { ok: true, id, inserted: true };
  }
});

export const listManualStoreRequests = query({
  args: {},
  handler: async (ctx) => {
    return await ctx.db
      .query("manual_store_requests")
      .withIndex("by_updated_at")
      .order("desc")
      .take(500);
  }
});

export const upsertManualStoreRequest = mutation({
  args: { payload: v.any() },
  handler: async (ctx, { payload }) => {
    const request_id = text(payload.request_id);
    const store_name = text(payload.store_name);
    if (!request_id || !store_name) throw new Error("request_id dan store_name wajib diisi");

    const now = Date.now();
    const doc = {
      request_id,
      status: text(payload.status, "pending"),
      created_at: num(payload.created_at, now),
      updated_at: num(payload.updated_at, now),
      bestie_name: text(payload.bestie_name),
      store_name,
      store_code: text(payload.store_code),
      address: text(payload.address),
      note: text(payload.note),
      session_id: text(payload.session_id),
      page_url: text(payload.page_url),
      user_agent: text(payload.user_agent)
    };

    const existing = await ctx.db
      .query("manual_store_requests")
      .withIndex("by_request_id", (q) => q.eq("request_id", request_id))
      .unique();

    if (existing) {
      await ctx.db.patch(existing._id, doc);
      return { ok: true, id: existing._id, updated: true };
    }

    const id = await ctx.db.insert("manual_store_requests", doc);
    return { ok: true, id, inserted: true };
  }
});

Optional HTTP actions are no longer required for realtime. The frontend now uses ConvexClient WebSocket subscriptions and only falls back to HTTP/polling when realtime is unavailable.
*/
