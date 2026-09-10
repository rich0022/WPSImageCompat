interface AssetBinding {
  fetch(request: Request): Promise<Response>;
}

interface Env {
  ASSETS: AssetBinding;
}

/** Keep update metadata out of edge and browser caches while leaving all other assets immutable. */
export function noStoreResponse(response: Response): Response {
  const headers = new Headers(response.headers);
  headers.set('Cache-Control', 'no-store, no-cache, max-age=0, must-revalidate');
  return new Response(response.body, { status: response.status, statusText: response.statusText, headers });
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    return noStoreResponse(await env.ASSETS.fetch(request));
  },
};
