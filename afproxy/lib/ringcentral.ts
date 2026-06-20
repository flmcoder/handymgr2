// @ts-nocheck - Deno code, type checking disabled
const RC_SERVER_URL = String(Deno.env.get("RC_SERVER_URL") || "").trim();
const RC_TOKEN = String(Deno.env.get("RC_ACCESS_TOKEN") || "").trim();
const RC_FROM = String(Deno.env.get("RC_FROM_NUMBER") || "").trim();

function normalizePhone(phone: string): string {
  const digits = String(phone || "").replace(/\D/g, "");
  if (digits.length === 10) return `+1${digits}`;
  if (digits.length === 11 && digits.startsWith("1")) return `+${digits}`;
  if (digits.length > 11) return `+${digits}`;
  return String(phone || "").trim();
}

export async function sendSMS(
  to: string,
  text: string,
): Promise<{ ok: boolean; id?: string; error?: string }> {
  const toNumber = normalizePhone(to);
  if (!toNumber) return { ok: false, error: "Missing recipient phone" };

  if (!RC_SERVER_URL || !RC_TOKEN || !RC_FROM) {
    console.warn("[ringcentral] env vars missing; SMS send skipped");
    return { ok: true, id: "mock-sms-no-env" };
  }

  try {
    const url = `${RC_SERVER_URL}/restapi/v1.0/account/~/extension/~/sms`;
    const resp = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${RC_TOKEN}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: { phoneNumber: RC_FROM },
        to: [{ phoneNumber: toNumber }],
        text,
      }),
    });

    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      return {
        ok: false,
        error: `RingCentral HTTP ${resp.status}: ${body.slice(0, 180)}`,
      };
    }

    const data = await resp.json().catch(() => ({}));
    return { ok: true, id: String(data?.id || "sms-sent") };
  } catch (err: any) {
    return { ok: false, error: String(err?.message || err || "SMS failed") };
  }
}
