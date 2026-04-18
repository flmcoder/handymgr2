// ============================================================================
// tests/billing_smoke.ts — manual route-action smoke test for billing endpoints.
//
// Usage:
//   API_PROXY="https://<your-proxy-url>" \
//   PROXY_BEARER_TOKEN="<optional-token>" \
//   BILLING_GROUP_ID="<group-uuid>" \
//   BILLING_VENDOR_ID="<vendor-id>" \
//   BILLING_PROPERTY_ID="<property-id>" \
//   BILLING_WO_ID="<wo-uuid>" \
//   BILLING_WO_NUMBER="<wo-number>" \
//   BILLING_INVOICE_NUMBER="<invoice-number>" \
//   deno run --allow-env --allow-net tests/billing_smoke.ts
//
// Notes:
// - This script calls the proxy action router directly over HTTP.
// - It is safe for read-only checks; all actions here are GET-style read actions.
// ============================================================================

type ActionCase = {
  action: string;
  params: Record<string, string>;
  required?: boolean;
};

const API_PROXY = String(Deno.env.get("API_PROXY") || "").trim();
const TOKEN = String(Deno.env.get("PROXY_BEARER_TOKEN") || "").trim();

if (!API_PROXY) {
  console.error("Missing API_PROXY env var.");
  Deno.exit(1);
}

function q(name: string): string {
  return String(Deno.env.get(name) || "").trim();
}

function buildUrl(action: string, params: Record<string, string>): string {
  const url = new URL(API_PROXY);
  url.searchParams.set("action", action);
  for (const [k, v] of Object.entries(params)) {
    if (v) url.searchParams.set(k, v);
  }
  return url.toString();
}

async function callAction(item: ActionCase): Promise<void> {
  const startedAt = performance.now();
  const url = buildUrl(item.action, item.params);
  const headers: HeadersInit = { Accept: "application/json" };
  if (TOKEN) headers.Authorization = `Bearer ${TOKEN}`;

  const res = await fetch(url, { headers });
  let payload: any = null;
  try {
    payload = await res.json();
  } catch {
    payload = null;
  }

  if (!res.ok) {
    const ms = Math.round(performance.now() - startedAt);
    console.log(`[${item.action}] HTTP ${res.status} ${res.statusText} (${ms}ms)`);
    return;
  }

  const ok = !!(payload && payload.ok !== false);
  const total = Number(payload?.total || 0) || 0;
  const rows = Array.isArray(payload?.data)
    ? payload.data.length
    : (Array.isArray(payload?.results) ? payload.results.length : 0);
  const err = payload?.error ? String(payload.error) : "";

  if (ok) {
    const ms = Math.round(performance.now() - startedAt);
    console.log(`[${item.action}] ok rows=${rows} total=${total} (${ms}ms)`);
  } else {
    const ms = Math.round(performance.now() - startedAt);
    console.log(`[${item.action}] not-ok error=${err || "unknown"} (${ms}ms)`);
  }
}

const today = new Date();
const yyyy = today.getFullYear();
const mm = String(today.getMonth() + 1).padStart(2, "0");
const dd = String(today.getDate()).padStart(2, "0");
const nowDate = `${yyyy}-${mm}-${dd}`;

const actions: ActionCase[] = [
  {
    action: "bills_list",
    params: { group_id: q("BILLING_GROUP_ID"), limit: "5", offset: "0" },
    required: true,
  },
  {
    action: "bills_by_vendor",
    params: { vendor_id: q("BILLING_VENDOR_ID"), group_id: q("BILLING_GROUP_ID"), limit: "5", offset: "0" },
  },
  {
    action: "bills_by_property",
    params: { property_id: q("BILLING_PROPERTY_ID"), group_id: q("BILLING_GROUP_ID"), limit: "5", offset: "0" },
  },
  {
    action: "bills_by_wo",
    params: { wo_id: q("BILLING_WO_ID"), limit: "5", offset: "0" },
  },
  {
    action: "bills_by_wo_number",
    params: { wo_number: q("BILLING_WO_NUMBER"), limit: "5", offset: "0" },
  },
  {
    action: "bills_by_invoice",
    params: { invoice_number: q("BILLING_INVOICE_NUMBER"), limit: "5", offset: "0" },
  },
  {
    action: "bills_due_range",
    params: {
      group_id: q("BILLING_GROUP_ID"),
      due_from: q("BILLING_DUE_FROM") || nowDate,
      due_to: q("BILLING_DUE_TO") || nowDate,
      limit: "5",
      offset: "0",
    },
    required: true,
  },
];

console.log("Billing smoke test start");
console.log(`Proxy: ${API_PROXY}`);

for (const item of actions) {
  const hasAllParams = Object.values(item.params).every((v) => String(v || "").trim() !== "");
  if (!hasAllParams && !item.required) {
    console.log(`[${item.action}] skipped (missing optional test params)`);
    continue;
  }
  await callAction(item);
}

console.log("Billing smoke test done");
