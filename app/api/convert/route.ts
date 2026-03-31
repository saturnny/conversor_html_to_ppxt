import { renderHtmlToSlides } from "../../../lib/server/html-slide-renderer";
import { buildPresentationFromRenderedSlides } from "../../../lib/server/pptx-builder";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const upload = formData.get("htmlFile");

    if (!(upload instanceof File)) {
      return Response.json(
        {
          error: "Envie um arquivo HTML valido para a conversao."
        },
        { status: 400 }
      );
    }

    if (!isHtmlFile(upload.name)) {
      return Response.json(
        {
          error: "O arquivo precisa ter extensao .html ou .htm."
        },
        { status: 400 }
      );
    }

    const html = await decodeHtmlFile(upload);
    const rendered = await renderHtmlToSlides(html, upload.name);

    if (!rendered.slides.length) {
      return Response.json(
        {
          error: "Nenhum slide foi detectado no HTML enviado."
        },
        { status: 422 }
      );
    }

    const buffer = await buildPresentationFromRenderedSlides(rendered);
    const fileName = sanitizeDownloadName(rendered.fileName || "apresentacao.pptx");

    return new Response(new Uint8Array(buffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="${fileName}"`,
        "Cache-Control": "no-store",
        "X-Slide-Count": String(rendered.slides.length),
        "X-Conversion-Warnings": JSON.stringify(rendered.warnings)
      }
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Nao foi possivel converter o HTML em PowerPoint.";

    return Response.json(
      {
        error: message
      },
      { status: 500 }
    );
  }
}

function sanitizeDownloadName(fileName: string): string {
  const cleaned = fileName.replace(/[^\w.\-]+/g, "-");
  return cleaned.toLowerCase().endsWith(".pptx") ? cleaned : `${cleaned}.pptx`;
}

function isHtmlFile(fileName: string): boolean {
  return /\.(html?|HTML?)$/.test(fileName);
}

async function decodeHtmlFile(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const utf8Text = new TextDecoder("utf-8").decode(bytes);
  const charset = sniffCharset(utf8Text);

  if (!charset || charset.toLowerCase() === "utf-8") {
    return utf8Text;
  }

  try {
    return new TextDecoder(charset).decode(bytes);
  } catch {
    return utf8Text;
  }
}

function sniffCharset(html: string): string | null {
  const charsetMatch = html.match(/<meta[^>]+charset=["']?\s*([a-z0-9\-_]+)/i);
  if (charsetMatch?.[1]) {
    return charsetMatch[1].trim();
  }

  const contentMatch = html.match(/content=["'][^"']*charset=([a-z0-9\-_]+)/i);
  return contentMatch?.[1]?.trim() ?? null;
}
