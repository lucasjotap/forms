import handler from "vinext/server/fetch-handler";
import { runWithConnectorBinding } from "../lib/connector-context";
import type { ConnectorBinding } from "../lib/connector-contract.mjs";

const PLATFORM_AUTH_PREFIX = "oai-authenticated-";

function isAdminRequest(request: Request, url: URL) {
  if (url.pathname === "/admin" || url.pathname.startsWith("/admin/")) return true;
  return url.pathname === "/api/responses" && (request.method === "GET" || request.method === "HEAD");
}

async function sameSecret(a: string, b: string) {
  const enc = new TextEncoder();
  const [x, y] = await Promise.all([crypto.subtle.digest("SHA-256", enc.encode(a)), crypto.subtle.digest("SHA-256", enc.encode(b))]);
  const ax = new Uint8Array(x), by = new Uint8Array(y);
  let diff = 0;
  for (let i = 0; i < ax.length; i++) diff |= ax[i] ^ by[i];
  return diff === 0;
}

async function basicAuthOk(request: Request, password: string) {
  const header = request.headers.get("authorization") || "";
  if (!header.startsWith("Basic ")) return false;
  let decoded = "";
  try { decoded = atob(header.slice(6)); } catch { return false; }
  const supplied = decoded.slice(decoded.indexOf(":") + 1);
  return sameSecret(supplied, password);
}

// Outside ChatGPT Sites nothing strips client-sent identity headers, so drop them and
// protect the organizer pages with a password stored as the ADMIN_PASSWORD secret.
async function applyAdminAuth(request: Request, env: Cloudflare.Env): Promise<Request | Response> {
  if (import.meta.env.DEV || env.TRUST_PLATFORM_AUTH === "1") return request;
  const headers = new Headers(request.headers);
  for (const name of [...headers.keys()]) if (name.startsWith(PLATFORM_AUTH_PREFIX)) headers.delete(name);
  const url = new URL(request.url);
  if (isAdminRequest(request, url)) {
    if (!env.ADMIN_PASSWORD || !env.ADMIN_EMAIL || !(await basicAuthOk(request, env.ADMIN_PASSWORD))) {
      return new Response("Acesso restrito ao organizador.", {
        status: 401,
        headers: { "WWW-Authenticate": 'Basic realm="Organizador", charset="UTF-8"', "Cache-Control": "no-store" },
      });
    }
    headers.set(PLATFORM_AUTH_PREFIX + "user-id", "organizer");
    headers.set(PLATFORM_AUTH_PREFIX + "user-email", env.ADMIN_EMAIL);
  }
  return new Request(request, { headers });
}

const worker = {
  async fetch(request: Request, env: Cloudflare.Env, ctx: ExecutionContext<{ CONNECTORS?: ConnectorBinding }>) {
    const authed = await applyAdminAuth(request, env);
    if (authed instanceof Response) return authed;
    request = authed;
    let binding = ctx.props?.CONNECTORS;
    // Local preview emulates the same request-scoped capability. This branch and
    // the auxiliary service binding are absent from production builds.
    if (import.meta.env.DEV && !binding && env.CONNECTORS) {
      const preview = env.CONNECTORS;
      const expiresAt = Date.now() + 60_000;
      binding = {
        async getContext() {
          if (Date.now() >= expiresAt) return { status: "request_context_expired" };
          return preview.getContext?.() ?? { status: "binding_unavailable" };
        },
        async invoke(connectorId, actionName, args) {
          if (Date.now() >= expiresAt) {
            return { status: "request_context_expired", message: "This request has expired. Please try again." };
          }
          return preview.invoke(connectorId, actionName, args);
        },
      };
    }
    return runWithConnectorBinding(binding, () => handler.fetch(request, env, ctx));
  },
};

export default worker;
