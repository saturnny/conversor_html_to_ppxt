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

  const bundle = patchBundleSource(await readFile(bundlePath, "utf-8"));

  return new Response(bundle, {
    headers: {
      "Content-Type": "application/javascript; charset=utf-8",
      "Cache-Control": "no-store"
    }
  });
}

function patchBundleSource(source: string): string {
  return source
    .replace(
      "lineSpacing = lhPx * 0.75 * scale;",
      "lineSpacing = Number((lhPx * 0.75 * scale).toFixed(2));"
    )
    .replace(
      "fontSize: Math.floor(fontSizePx * 0.75 * scale),",
      "fontSize: Number((fontSizePx * 0.75 * scale).toFixed(2)), charSpacing: !isNaN(parseFloat(style.letterSpacing)) ? Number((parseFloat(style.letterSpacing) * 0.75 * scale).toFixed(2)) : 0,"
    )
    .replace(
      "!isBgClipText && style.backgroundImage && style.backgroundImage.includes('linear-gradient');",
      "!isBgClipText && style.backgroundImage && style.backgroundImage.includes('linear-gradient') && !style.backgroundImage.includes('url(');"
    )
    .replaceAll(
      "Math.floor(textPayload.text[0]?.options?.fontSize) || 12",
      "Number((textPayload.text[0]?.options?.fontSize ?? 12).toFixed(2)) || 12"
    );
}
