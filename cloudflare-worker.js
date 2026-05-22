/**
 * Cloudflare Worker — Zammad CORS proxy for Q Check Outlook add-in
 *
 * Deploy this script at dash.cloudflare.com → Workers & Pages → Create Worker.
 * Copy this entire file into the editor and click "Deploy".
 *
 * After deploying, copy the Worker URL (e.g. https://qcheck-zammad.yourname.workers.dev)
 * and paste it into the "Zammad Proxy URL" field in the Q Check add-in Settings.
 *
 * The Worker:
 *   GET  /groups  → forwards to Zammad GET /api/v1/groups  (used by "Test APIs")
 *   POST /tickets → forwards to Zammad POST /api/v1/tickets (used by "Send to Maritime Team")
 *
 * Authentication: pass your personal Zammad token in the X-Zammad-Token header.
 * The Worker forwards it as: Authorization: Token token=<your-token>
 */

const ZAMMAD_BASE    = "https://euromar.zammad.com";
const ALLOWED_ORIGIN = "https://duropoint.github.io";

export default {
  async fetch(request) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }

    const url   = new URL(request.url);
    const token = request.headers.get("X-Zammad-Token");

    if (!token) {
      return jsonResponse({ error: "Missing X-Zammad-Token header" }, 401);
    }

    let zammadPath;
    if (url.pathname === "/tickets" && request.method === "POST") {
      zammadPath = "/api/v1/tickets";
    } else if (url.pathname === "/groups" && request.method === "GET") {
      zammadPath = "/api/v1/groups";
    } else {
      return jsonResponse({ error: "Not found" }, 404);
    }

    let upstream;
    try {
      upstream = await fetch(`${ZAMMAD_BASE}${zammadPath}`, {
        method:  request.method,
        headers: {
          "Authorization": `Token token=${token}`,
          "Content-Type":  "application/json",
        },
        body: request.method === "POST" ? request.body : null,
      });
    } catch (err) {
      return jsonResponse({ error: "Failed to reach Zammad: " + err.message }, 502);
    }

    const body = await upstream.text();
    return new Response(body, {
      status:  upstream.status,
      headers: {
        ...corsHeaders(),
        "Content-Type": upstream.headers.get("Content-Type") || "application/json",
      },
    });
  }
};

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin":  ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-Zammad-Token",
  };
}

function jsonResponse(obj, status) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { ...corsHeaders(), "Content-Type": "application/json" },
  });
}
