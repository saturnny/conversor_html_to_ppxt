import PptxGenJS from "pptxgenjs";
import type {
  BorderToken,
  ColorToken,
  ConversionManifest,
  RenderedPresentation,
  SlideElement,
  TableElement,
  TextElement,
  TextRun
} from "../presentation-types";

const PPT_WIDTH_IN = 13.333;
const PPT_HEIGHT_IN = 7.5;
const DEFAULT_FONT = "Aptos";

export async function buildPresentationFromRenderedSlides(
  presentation: RenderedPresentation
): Promise<Buffer> {
  const pptx = createPptxDocument(presentation.sourceName);

  for (const renderedSlide of presentation.slides) {
    const slide = pptx.addSlide();
    slide.background = {
      color: "FFFFFF"
    };

    slide.addImage({
      data: renderedSlide.imageDataUrl,
      x: 0,
      y: 0,
      w: PPT_WIDTH_IN,
      h: PPT_HEIGHT_IN
    });
  }

  const buffer = await pptx.write({ outputType: "nodebuffer" });
  return Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer as ArrayBuffer);
}

export async function buildPresentationBuffer(manifest: ConversionManifest): Promise<Buffer> {
  const pptx = createPptxDocument(manifest.sourceName);

  for (const slideManifest of manifest.slides) {
    const slide = pptx.addSlide();
    const { scale, offsetX, offsetY } = fitSlide(slideManifest.width, slideManifest.height);

    if (slideManifest.background) {
      slide.background = {
        color: stripHash(slideManifest.background.hex),
        transparency: alphaToTransparency(slideManifest.background.alpha)
      };
    }

    const sortedElements = [...slideManifest.elements].sort((left, right) => {
      if (left.zIndex === right.zIndex) {
        return left.order - right.order;
      }
      return left.zIndex - right.zIndex;
    });

    for (const element of sortedElements) {
      await renderElement(slide, pptx, element, scale, offsetX, offsetY);
    }
  }

  const buffer = await pptx.write({ outputType: "nodebuffer" });
  return Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer as ArrayBuffer);
}

function createPptxDocument(title: string) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Codex";
  pptx.company = "OpenAI";
  pptx.subject = "Conversao de HTML para PowerPoint";
  pptx.title = title;
  return pptx;
}

async function renderElement(
  slide: PptxGenJS.Slide,
  pptx: PptxGenJS,
  element: SlideElement,
  scale: number,
  offsetX: number,
  offsetY: number
): Promise<void> {
  const x = offsetX + element.x * scale;
  const y = offsetY + element.y * scale;
  const w = element.w * scale;
  const h = element.h * scale;

  if (w <= 0 || h <= 0) {
    return;
  }

  if (element.type === "shape") {
    slide.addShape(element.kind === "roundRect" ? pptx.ShapeType.roundRect : pptx.ShapeType.rect, {
      x,
      y,
      w,
      h,
      fill: toFillOptions(element.fill),
      line: toLineOptions(element.border, scale)
    });
    return;
  }

  if (element.type === "image") {
    const imageData = await resolveImageData(element.src);
    if (!imageData) {
      return;
    }

    slide.addImage({
      data: imageData,
      x,
      y,
      w,
      h
    });
    return;
  }

  if (element.type === "table") {
    renderTable(slide, element, scale, offsetX, offsetY);
    return;
  }

  renderText(slide, element, scale, offsetX, offsetY);
}

function renderText(
  slide: PptxGenJS.Slide,
  element: TextElement,
  scale: number,
  offsetX: number,
  offsetY: number
): void {
  const x = offsetX + element.x * scale;
  const y = offsetY + element.y * scale;
  const w = element.w * scale;
  const h = element.h * scale;

  const baseOptions = {
    x,
    y,
    w,
    h,
    margin: 0,
    fit: "shrink" as const,
    fontFace: element.fontFamily || DEFAULT_FONT,
    fontSize: pxToPoints(element.fontSizePx, scale),
    color: element.color ? stripHash(element.color.hex) : "1A1A1A",
    bold: element.bold,
    italic: element.italic,
    underline: toUnderlineOption(element.underline, element.color),
    align: element.align,
    valign: element.verticalAlign
  };

  const runs = element.runs
    .filter((run) => run.text.trim().length > 0)
    .map((run) => ({
      text: run.text,
      options: buildRunOptions(run, element, scale)
    }));

  if (!runs.length) {
    slide.addText(element.text, baseOptions);
    return;
  }

  slide.addText(runs, baseOptions);
}

function renderTable(
  slide: PptxGenJS.Slide,
  element: TableElement,
  scale: number,
  offsetX: number,
  offsetY: number
): void {
  const rows = element.rows.map((row) =>
    row.map((cell) => ({
      text: cell.text,
      options: {
        fill: cell.background ? stripHash(cell.background.hex) : undefined,
        color: cell.color ? stripHash(cell.color.hex) : "2B241D",
        fontFace: cell.fontFamily || DEFAULT_FONT,
        fontSize: pxToPoints(cell.fontSizePx ?? 14, scale),
        bold: cell.bold,
        align: cell.align,
        valign: "mid",
        margin: 2
      }
    }))
  );

  slide.addTable(rows as never, {
    x: offsetX + element.x * scale,
    y: offsetY + element.y * scale,
    w: element.w * scale,
    colW: element.colWidths.map((width) => width * scale),
    rowH: element.rowHeights.map((height) => Math.max(height * scale, 0.2)),
    border: element.border
      ? {
          type: "solid",
          color: stripHash(element.border.color.hex),
          pt: pxToPoints(element.border.widthPx, scale)
        }
      : {
          type: "solid",
          color: "D8CBBF",
          pt: 0.75
        },
    fill: element.fill ? stripHash(element.fill.hex) : undefined,
    margin: 2
  } as never);
}

function buildRunOptions(run: TextRun, element: TextElement, scale: number) {
  const color = run.color || element.color || { hex: "#1A1A1A" };

  return {
    fontFace: run.fontFamily || element.fontFamily || DEFAULT_FONT,
    fontSize: pxToPoints(run.fontSizePx ?? element.fontSizePx, scale),
    color: stripHash(color.hex),
    bold: run.bold ?? element.bold,
    italic: run.italic ?? element.italic,
    underline: toUnderlineOption(run.underline ?? element.underline, color)
  };
}

function fitSlide(widthPx: number, heightPx: number) {
  const scale = Math.min(PPT_WIDTH_IN / widthPx, PPT_HEIGHT_IN / heightPx);
  const offsetX = (PPT_WIDTH_IN - widthPx * scale) / 2;
  const offsetY = (PPT_HEIGHT_IN - heightPx * scale) / 2;

  return {
    scale,
    offsetX,
    offsetY
  };
}

function pxToPoints(px: number, scale: number): number {
  return Math.max(px * scale * 72, 6);
}

function toFillOptions(color: ColorToken | null | undefined) {
  if (!color) {
    return {
      color: "FFFFFF",
      transparency: 100
    };
  }

  return {
    color: stripHash(color.hex),
    transparency: alphaToTransparency(color.alpha)
  };
}

function toLineOptions(border: BorderToken | null | undefined, scale: number) {
  if (!border) {
    return {
      color: "FFFFFF",
      transparency: 100
    };
  }

  return {
    color: stripHash(border.color.hex),
    transparency: alphaToTransparency(border.color.alpha),
    pt: Math.max(pxToPoints(border.widthPx, scale), 0.5)
  };
}

function stripHash(hex: string): string {
  return hex.replace("#", "");
}

function alphaToTransparency(alpha = 1): number {
  return Math.round((1 - alpha) * 100);
}

function toUnderlineOption(enabled: boolean | undefined, color: ColorToken | null | undefined) {
  if (!enabled) {
    return undefined;
  }

  return {
    color: stripHash((color || { hex: "#1A1A1A" }).hex)
  };
}

async function resolveImageData(source: string): Promise<string | null> {
  if (source.startsWith("data:image")) {
    return source;
  }

  if (!/^https?:\/\//i.test(source)) {
    return null;
  }

  const response = await fetch(source, {
    signal: AbortSignal.timeout(15000)
  });

  if (!response.ok) {
    return null;
  }

  const contentType = response.headers.get("content-type") || "image/png";
  const buffer = Buffer.from(await response.arrayBuffer());
  return `data:${contentType};base64,${buffer.toString("base64")}`;
}
