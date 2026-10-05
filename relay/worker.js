/**
 * Cloudflare Worker: relay between Meta's WhatsApp webhook and the Google Apps Script web app.
 *
 * Why it exists: Meta requires the webhook to answer with a plain 200 (and to echo a challenge
 * during setup). Apps Script web apps always answer with a redirect, which Meta rejects.
 * This Worker answers Meta immediately, checks Meta's signature, and forwards real messages.
 *
 * Secrets (Worker → Settings → Variables and Secrets):
 *   VERIFY_TOKEN     any phrase you choose; type the same phrase into Meta's webhook setup
 *   APP_SECRET       Meta app → App settings → Basic → App secret
 *   APPS_SCRIPT_URL  the Apps Script web app URL (ends in /exec)
 *   RELAY_KEY        same long random string as the RELAY_KEY Script Property
 */
export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // One-time webhook verification from Meta
    if (request.method === "GET") {
      if (url.searchParams.get("hub.mode") === "subscribe" &&
          url.searchParams.get("hub.verify_token") === env.VERIFY_TOKEN) {
        return new Response(url.searchParams.get("hub.challenge") || "", { status: 200 });
      }
      return new Response("forbidden", { status: 403 });
    }

    if (request.method !== "POST") return new Response("method not allowed", { status: 405 });

    const raw = await request.arrayBuffer();
    const ok = await validSignature(raw, request.headers.get("x-hub-signature-256") || "", env.APP_SECRET);
    if (!ok) return new Response("bad signature", { status: 401 });

    const text = new TextDecoder().decode(raw);
    let hasMessages = false;
    try {
      const body = JSON.parse(text);
      hasMessages = (body.entry || []).some(e => (e.changes || []).some(c => c.value && Array.isArray(c.value.messages)));
    } catch (_) { /* not JSON: ignore */ }

    // Delivery/read receipts are acknowledged and dropped; only real messages go to the sheet.
    if (hasMessages) {
      ctx.waitUntil(
        fetch(env.APPS_SCRIPT_URL + "?key=" + encodeURIComponent(env.RELAY_KEY), {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: text,
          redirect: "follow",
        }).catch(err => console.error("forward failed", err))
      );
    }
    return new Response("ok", { status: 200 });
  },
};

async function validSignature(raw, header, secret) {
  if (!secret || !header.startsWith("sha256=")) return false;
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const mac = new Uint8Array(await crypto.subtle.sign("HMAC", key, raw));
  const expected = [...mac].map(b => b.toString(16).padStart(2, "0")).join("");
  const given = header.slice(7).toLowerCase();
  if (given.length !== expected.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ given.charCodeAt(i);
  return diff === 0;
}
