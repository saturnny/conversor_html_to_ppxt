import { readFile } from "node:fs/promises";
import path from "node:path";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  const bundlePath = path.join(
    process.cwd(),
    "node_modules",
    "dom-to-pptx",
    "dist",
    "dom-to-pptx.bundle.js"
  );

  const bundle = await readFile(bundlePath, "utf-8");

  return new Response(bundle, {
    headers: {
      "Content-Type": "application/javascript; charset=utf-8",
      "Cache-Control": "public, max-age=3600"
    }
  });
}
