import http from "node:http";
import https from "node:https";
import { isIP } from "node:net";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function GET(request: Request) {
  const requestUrl = new URL(request.url);
  const targetUrl = requestUrl.searchParams.get("url");

  if (!targetUrl) {
    return Response.json(
      {
        error: "Informe a URL do asset remoto."
      },
      { status: 400 }
    );
  }

  let parsed: URL;

  try {
    parsed = new URL(targetUrl);
  } catch {
    return Response.json(
      {
        error: "A URL informada eh invalida."
      },
      { status: 400 }
    );
  }

  if (!isAllowedRemoteUrl(parsed)) {
    return Response.json(
      {
        error: "A URL informada nao pode ser acessada por seguranca."
      },
      { status: 403 }
    );
  }

  try {
    const asset = await requestRemoteAsset(parsed, 0);
    const headers = new Headers();
    headers.set("Content-Type", asset.contentType || "application/octet-stream");
    headers.set("Cache-Control", "public, max-age=3600");
    headers.set("Access-Control-Allow-Origin", "*");

    return new Response(new Uint8Array(asset.buffer), {
      status: 200,
      headers
    });
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "Nao foi possivel carregar o asset remoto.";

    return Response.json(
      {
        error: message
      },
      { status: 502 }
    );
  }
}

type RemoteAssetResult = {
  buffer: Buffer;
  contentType: string | null;
};

async function requestRemoteAsset(url: URL, redirectCount: number): Promise<RemoteAssetResult> {
  if (redirectCount > 5) {
    throw new Error("O asset remoto excedeu o limite de redirecionamentos.");
  }

  const transport = url.protocol === "https:" ? https : http;

  return new Promise((resolve, reject) => {
    const request = transport.request(
      url,
      {
        method: "GET",
        headers: {
          "User-Agent": "html-to-pptx-app/1.0"
        },
        ...(url.protocol === "https:" ? { rejectUnauthorized: false } : {})
      },
      (response) => {
        const statusCode = response.statusCode || 0;

        if ([301, 302, 303, 307, 308].includes(statusCode) && response.headers.location) {
          response.resume();

          try {
            const redirectedUrl = new URL(response.headers.location, url);
            if (!isAllowedRemoteUrl(redirectedUrl)) {
              reject(new Error("O redirecionamento do asset remoto nao eh permitido."));
              return;
            }

            requestRemoteAsset(redirectedUrl, redirectCount + 1).then(resolve).catch(reject);
          } catch {
            reject(new Error("O asset remoto retornou um redirecionamento invalido."));
          }

          return;
        }

        if (statusCode < 200 || statusCode >= 300) {
          response.resume();
          reject(new Error(`Nao foi possivel carregar o asset remoto (${statusCode}).`));
          return;
        }

        const chunks: Buffer[] = [];

        response.on("data", (chunk) => {
          chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
        });

        response.on("end", () => {
          resolve({
            buffer: Buffer.concat(chunks),
            contentType: response.headers["content-type"] || null
          });
        });

        response.on("error", reject);
      }
    );

    request.setTimeout(30000, () => {
      request.destroy(new Error("Tempo esgotado ao buscar o asset remoto."));
    });

    request.on("error", reject);
    request.end();
  });
}

function isAllowedRemoteUrl(url: URL): boolean {
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    return false;
  }

  const hostname = url.hostname.trim().toLowerCase();
  if (!hostname) {
    return false;
  }

  if (
    hostname === "localhost" ||
    hostname === "0.0.0.0" ||
    hostname.endsWith(".local") ||
    hostname.endsWith(".internal")
  ) {
    return false;
  }

  if (!isIP(hostname)) {
    return true;
  }

  return !isPrivateIp(hostname);
}

function isPrivateIp(value: string): boolean {
  if (value.includes(":")) {
    const normalized = value.toLowerCase();
    return (
      normalized === "::1" ||
      normalized.startsWith("fc") ||
      normalized.startsWith("fd") ||
      normalized.startsWith("fe80:")
    );
  }

  const octets = value.split(".").map((part) => Number(part));
  if (octets.length !== 4 || octets.some((part) => Number.isNaN(part))) {
    return true;
  }

  if (octets[0] === 10 || octets[0] === 127) {
    return true;
  }

  if (octets[0] === 169 && octets[1] === 254) {
    return true;
  }

  if (octets[0] === 172 && octets[1] >= 16 && octets[1] <= 31) {
    return true;
  }

  if (octets[0] === 192 && octets[1] === 168) {
    return true;
  }

  return false;
}
