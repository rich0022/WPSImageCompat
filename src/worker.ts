interface AssetBinding {
  fetch(request: Request): Promise<Response>;
}

interface Env {
  ASSETS: AssetBinding;
}

// Keep the public sideload manifest on a unique asset URL whenever its ribbon changes.
// This avoids a stale edge copy keeping an older command layout after a deployment.
const manifestRevision = '1.5.1.0';

/** Keep update metadata out of edge and browser caches while leaving all other assets immutable. */
export function noStoreResponse(response: Response): Response {
  const headers = new Headers(response.headers);
  headers.set('Cache-Control', 'no-store, no-cache, max-age=0, must-revalidate');
  return new Response(response.body, { status: response.status, statusText: response.statusText, headers });
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname === '/manifest.xml') url.searchParams.set('__manifestRevision', manifestRevision);
    return noStoreResponse(await env.ASSETS.fetch(url.href === request.url ? request : new Request(url, request)));
  },
};
