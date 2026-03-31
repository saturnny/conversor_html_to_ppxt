import { toPng } from "html-to-image";
import type {
  BorderToken,
  ColorToken,
  ConversionManifest,
  HorizontalAlign,
  ImageElement,
  ShapeElement,
  SlideElement,
  SlideManifest,
  TableCellManifest,
  TableElement,
  TextElement,
  VerticalAlign
} from "../presentation-types";

const RENDER_WIDTH = 1280;
const MIN_SLIDE_HEIGHT = 720;
const INLINE_TAGS = new Set([
  "A",
  "ABBR",
  "B",
  "BR",
  "CODE",
  "EM",
  "I",
  "MARK",
  "S",
  "SMALL",
  "SPAN",
  "STRONG",
  "SUB",
  "SUP",
  "U"
]);

type BuildStatus = "reading" | "rendering";

type BuildStatusCallback = (status: BuildStatus) => void;

type SlideSource = {
  id: string;
  title: string;
  bounds: Bounds;
  roots: HTMLElement[];
  background?: ColorToken | null;
  frameRoot?: HTMLElement | null;
  snapshotRoot?: HTMLElement | null;
  prepare?: () => (() => void) | void;
};

type Bounds = {
  left: number;
  top: number;
  width: number;
  height: number;
};

type ElementCollector = {
  elements: SlideElement[];
  warnings: string[];
  push: (element: SlideElement) => void;
};

export async function buildManifestFromFile(
  file: File,
  onStatus?: BuildStatusCallback
): Promise<ConversionManifest> {
  onStatus?.("reading");
  const html = await decodeHtmlFile(file);
  onStatus?.("rendering");
  return buildManifestFromHtml(html, file.name);
}

async function buildManifestFromHtml(html: string, sourceName: string): Promise<ConversionManifest> {
  const iframe = document.createElement("iframe");
  iframe.setAttribute("sandbox", "allow-same-origin");
  iframe.style.position = "fixed";
  iframe.style.left = "-20000px";
  iframe.style.top = "0";
  iframe.style.width = `${RENDER_WIDTH}px`;
  iframe.style.height = `${MIN_SLIDE_HEIGHT}px`;
  iframe.style.opacity = "0";
  iframe.style.pointerEvents = "none";
  iframe.style.border = "0";

  document.body.appendChild(iframe);

  try {
    const renderableHtml = extractRenderableHtml(html);
    await loadIframeDocument(iframe, renderableHtml);

    const doc = iframe.contentDocument;
    if (!doc?.body) {
      throw new Error("Nao foi possivel renderizar o HTML enviado.");
    }

    await waitForDocumentReadiness(doc);

    const slides = await extractSlides(doc);
    if (!slides.length) {
      throw new Error("Nenhum conteudo visivel foi encontrado no HTML enviado.");
    }

    const warnings = dedupeMessages(slides.flatMap((slide) => slide.warnings));

    return {
      fileName: toOutputFileName(sourceName),
      sourceName,
      slides,
      warnings
    };
  } finally {
    iframe.remove();
  }
}

async function decodeHtmlFile(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const utf8Text = new TextDecoder("utf-8").decode(bytes);
  const sniffedCharset = sniffCharset(utf8Text);

  if (!sniffedCharset || sniffedCharset.toLowerCase() === "utf-8") {
    return utf8Text;
  }

  try {
    return new TextDecoder(sniffedCharset).decode(bytes);
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

async function loadIframeDocument(iframe: HTMLIFrameElement, html: string): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    iframe.onload = () => resolve();
    iframe.onerror = () => reject(new Error("Falha ao carregar o HTML no ambiente de renderizacao."));
    iframe.srcdoc = html;
  });
}

async function waitForDocumentReadiness(doc: Document): Promise<void> {
  const imagePromises = Array.from(doc.images).map(
    (image) =>
      new Promise<void>((resolve) => {
        if (image.complete) {
          resolve();
          return;
        }

        const finish = () => resolve();
        image.addEventListener("load", finish, { once: true });
        image.addEventListener("error", finish, { once: true });
        window.setTimeout(finish, 3000);
      })
  );

  await Promise.allSettled(imagePromises);
  await doc.fonts?.ready;
  await waitForAnimationFrame();
  await waitForAnimationFrame();
}

function waitForAnimationFrame(): Promise<void> {
  return new Promise((resolve) => {
    window.requestAnimationFrame(() => resolve());
  });
}

async function extractSlides(doc: Document): Promise<SlideManifest[]> {
  const pageWidth = Math.max(RENDER_WIDTH, doc.documentElement.scrollWidth, doc.body.scrollWidth);
  const pageHeight = Math.max(MIN_SLIDE_HEIGHT, doc.documentElement.scrollHeight, doc.body.scrollHeight);
  const pageBackground =
    toColorToken(getComputedStyle(doc.body).backgroundColor) ??
    toColorToken(getComputedStyle(doc.documentElement).backgroundColor) ?? {
      hex: "#FFFFFF",
      alpha: 1
    };

  const slideSources = detectSlideSources(doc, pageWidth, pageHeight, pageBackground);

  return Promise.all(
    slideSources.map(async (source) => {
    const cleanup = source.prepare?.();
    let orderCounter = 0;
    const collector: ElementCollector = {
      elements: [],
      warnings: [],
      push(element) {
        collector.elements.push(element);
      }
    };

    try {
      if (source.snapshotRoot) {
        const snapshotElement = await captureSlideSnapshot(source.snapshotRoot, source.bounds, collector.warnings);

        if (snapshotElement) {
          collector.push(snapshotElement);
        } else {
          for (const root of source.roots) {
            traverseElement(
              root,
              source.bounds,
              collector,
              () => orderCounter++,
              root === doc.body || root === source.frameRoot
            );
          }
        }
      } else {
        for (const root of source.roots) {
          traverseElement(
            root,
            source.bounds,
            collector,
            () => orderCounter++,
            root === doc.body || root === source.frameRoot
          );
        }
      }
    } finally {
      cleanup?.();
    }

    const elements = collector.elements
      .filter((element) => element.w > 0 && element.h > 0)
      .sort((left, right) => {
        if (left.zIndex === right.zIndex) {
          return left.order - right.order;
        }
        return left.zIndex - right.zIndex;
      });

    return {
      id: source.id,
      title: source.title,
      width: source.bounds.width,
      height: source.bounds.height,
      background: source.background ?? pageBackground,
      elements,
      warnings: dedupeMessages(collector.warnings)
    };
    })
  );
}

function detectSlideSources(
  doc: Document,
  pageWidth: number,
  pageHeight: number,
  pageBackground: ColorToken
): SlideSource[] {
  const body = doc.body;
  const explicitRoots = findExplicitSlideRoots(doc);

  if (explicitRoots.length > 1) {
    return explicitRoots.map((root, index) => createExplicitSlideSource(root, index + 1, explicitRoots.length, pageBackground));
  }

  const directChildren = Array.from(body.children).filter(
    (node): node is HTMLElement => node instanceof HTMLElement && isVisibleElement(node)
  );

  const semanticRoots = directChildren.filter((element) =>
    /^(SECTION|ARTICLE|MAIN|HEADER|FOOTER)$/.test(element.tagName)
  );

  if (semanticRoots.length > 1) {
    return semanticRoots.map((root, index) => createSlideSource([root], index + 1, pageWidth));
  }

  const groupedSources = groupChildrenIntoSlides(directChildren, pageWidth);
  if (groupedSources.length > 1) {
    return groupedSources;
  }

  return [
    {
      id: "slide-1",
      title: extractTitleFromDocument(doc),
      bounds: {
        left: 0,
        top: 0,
        width: pageWidth,
        height: pageHeight
      },
      roots: [body]
    }
  ];
}

function groupChildrenIntoSlides(children: HTMLElement[], pageWidth: number): SlideSource[] {
  if (children.length < 2) {
    return [];
  }

  const groups: HTMLElement[][] = [];
  let currentGroup: HTMLElement[] = [];
  let groupTop = 0;

  for (const child of children) {
    const rect = getAbsoluteRect(child);

    if (!currentGroup.length) {
      currentGroup = [child];
      groupTop = rect.top;
      continue;
    }

    const projectedHeight = rect.bottom - groupTop;
    const shouldBreak = projectedHeight > MIN_SLIDE_HEIGHT * 1.15;

    if (shouldBreak) {
      groups.push(currentGroup);
      currentGroup = [child];
      groupTop = rect.top;
      continue;
    }

    currentGroup.push(child);
  }

  if (currentGroup.length) {
    groups.push(currentGroup);
  }

  if (groups.length <= 1) {
    return [];
  }

  return groups.map((group, index) => createSlideSource(group, index + 1, pageWidth));
}

function findExplicitSlideRoots(doc: Document): HTMLElement[] {
  const candidates = Array.from(
    doc.querySelectorAll(
      ".presentation-container > .slide, body > .slide, [data-slide], [data-ppt-slide], .ppt-slide, .page-slide"
    )
  ).filter((node): node is HTMLElement => node instanceof HTMLElement);

  const uniqueCandidates = Array.from(new Set(candidates));

  return uniqueCandidates.filter((element) => {
    const rect = getAbsoluteRect(element);
    return rect.width > 160 && rect.height > 120;
  });
}

function createExplicitSlideSource(
  root: HTMLElement,
  index: number,
  totalSlides: number,
  pageBackground: ColorToken
): SlideSource {
  const cleanup = prepareExplicitSlide(root, index, totalSlides);

  try {
    const rect = getAbsoluteRect(root);
    return {
      id: `slide-${index}`,
      title: root.dataset.slideTitle?.trim() || inferSlideTitle(root, index),
      bounds: {
        left: rect.left,
        top: rect.top,
        width: rect.width,
        height: rect.height
      },
      roots: [root],
      frameRoot: root,
      snapshotRoot: root,
      background: toColorToken(getComputedStyle(root).backgroundColor) ?? pageBackground,
      prepare: () => prepareExplicitSlide(root, index, totalSlides)
    };
  } finally {
    cleanup();
  }
}

function createSlideSource(roots: HTMLElement[], index: number, pageWidth: number): SlideSource {
  const rects = roots.map(getAbsoluteRect);
  const top = Math.min(...rects.map((rect) => rect.top));
  const bottom = Math.max(...rects.map((rect) => rect.bottom));
  const title = roots[0]?.dataset.slideTitle?.trim() || inferSlideTitle(roots[0], index);

  return {
    id: `slide-${index}`,
    title,
    bounds: {
      left: 0,
      top,
      width: pageWidth,
      height: Math.max(MIN_SLIDE_HEIGHT, bottom - top)
    },
    roots
  };
}

function prepareExplicitSlide(root: HTMLElement, index: number, totalSlides: number): () => void {
  const computedDisplay = getComputedStyle(root).display;
  const previousState = {
    visibility: readInlineStyle(root, "visibility"),
    opacity: readInlineStyle(root, "opacity"),
    display: readInlineStyle(root, "display"),
    zIndex: readInlineStyle(root, "z-index")
  };
  const hadActiveClass = root.classList.contains("active");

  root.classList.add("active");
  root.style.setProperty("visibility", "visible", "important");
  root.style.setProperty("opacity", "1", "important");
  root.style.setProperty("display", computedDisplay === "none" ? "flex" : computedDisplay || "block", "important");
  root.style.setProperty("z-index", "999", "important");

  const indicatorState = Array.from(root.querySelectorAll(".page-indicator")).map((element) => {
    const target = element as HTMLElement;
    const previousText = target.textContent ?? "";
    target.textContent = `${index} / ${totalSlides}`;
    return () => {
      target.textContent = previousText;
    };
  });

  const progressState = Array.from(root.querySelectorAll(".footer-progress-bar")).map((element) => {
    const target = element as HTMLElement;
    const previousWidth = readInlineStyle(target, "width");
    target.style.setProperty("width", `${(index / totalSlides) * 100}%`, "important");
    return () => {
      restoreInlineStyle(target, "width", previousWidth);
    };
  });

  return () => {
    if (!hadActiveClass) {
      root.classList.remove("active");
    }

    restoreInlineStyle(root, "visibility", previousState.visibility);
    restoreInlineStyle(root, "opacity", previousState.opacity);
    restoreInlineStyle(root, "display", previousState.display);
    restoreInlineStyle(root, "z-index", previousState.zIndex);

    indicatorState.forEach((restore) => restore());
    progressState.forEach((restore) => restore());
  };
}

async function captureSlideSnapshot(
  root: HTMLElement,
  bounds: Bounds,
  warnings: string[]
): Promise<ImageElement | null> {
  await waitForAnimationFrame();
  await waitForAnimationFrame();

  try {
    const computedBackground = getComputedStyle(root).backgroundColor;
    const dataUrl = await toPng(root, {
      cacheBust: true,
      pixelRatio: 2,
      skipFonts: false,
      backgroundColor:
        computedBackground && computedBackground !== "rgba(0, 0, 0, 0)" ? computedBackground : "#FFFFFF",
      width: Math.max(1, Math.round(bounds.width)),
      height: Math.max(1, Math.round(bounds.height))
    });

    return {
      type: "image",
      x: 0,
      y: 0,
      w: bounds.width,
      h: bounds.height,
      src: dataUrl,
      alt: "Slide renderizado a partir do HTML",
      order: 0,
      zIndex: -1000
    };
  } catch {
    warnings.push(
      "Nao foi possivel capturar o slide como imagem renderizada. Foi usado o modo estruturado como fallback."
    );
    return null;
  }
}

function extractRenderableHtml(html: string): string {
  const normalized = unwrapEmbeddedHtmlDocument(html).trim();
  return normalized || html;
}

function unwrapEmbeddedHtmlDocument(html: string): string {
  const lower = html.toLowerCase();
  const doctypeMatches = Array.from(html.matchAll(/<!doctype html>/gi));

  if (doctypeMatches.length > 1 && doctypeMatches[1].index !== undefined) {
    const start = doctypeMatches[1].index;
    const end = lower.indexOf("</html>", start);
    if (end !== -1) {
      return html.slice(start, end + "</html>".length);
    }
    return html.slice(start);
  }

  const htmlMatches = Array.from(html.matchAll(/<html[\s>]/gi));
  if (htmlMatches.length > 1 && htmlMatches[1].index !== undefined) {
    const start = htmlMatches[1].index;
    const end = lower.indexOf("</html>", start);
    if (end !== -1) {
      return html.slice(start, end + "</html>".length);
    }
    return html.slice(start);
  }

  return html;
}

function readInlineStyle(element: HTMLElement, property: string) {
  return {
    value: element.style.getPropertyValue(property),
    priority: element.style.getPropertyPriority(property)
  };
}

function restoreInlineStyle(
  element: HTMLElement,
  property: string,
  previous: { value: string; priority: string }
): void {
  if (previous.value) {
    element.style.setProperty(property, previous.value, previous.priority);
    return;
  }

  element.style.removeProperty(property);
}

function traverseElement(
  element: HTMLElement,
  bounds: Bounds,
  collector: ElementCollector,
  nextOrder: () => number,
  isRoot = false
): void {
  if (!isVisibleElement(element)) {
    return;
  }

  const rect = getAbsoluteRect(element);
  if (!intersectsBounds(rect, bounds)) {
    return;
  }

  const style = getComputedStyle(element);
  const tag = element.tagName.toLowerCase();
  const zIndex = parseZIndex(style.zIndex);
  const childElements = Array.from(element.children).filter((child): child is HTMLElement => child instanceof HTMLElement);

  if (tag === "img" || tag === "svg" || tag === "canvas") {
    const imageElement = buildMediaElement(element, bounds, collector.warnings, nextOrder(), zIndex);
    if (imageElement) {
      collector.push(imageElement);
    }
    return;
  }

  if (tag === "table") {
    const tableElement = buildTableElement(element as HTMLTableElement, bounds, nextOrder(), zIndex);
    if (tableElement) {
      collector.push(tableElement);
    }
    return;
  }

  if (tag === "ul" || tag === "ol") {
    buildListElements(element as HTMLUListElement | HTMLOListElement, bounds, collector, nextOrder, zIndex);
    return;
  }

  if (tag === "hr") {
    collector.push({
      type: "shape",
      kind: "rect",
      x: rect.left - bounds.left,
      y: rect.top - bounds.top,
      w: rect.width,
      h: Math.max(rect.height, 1),
      fill: toColorToken(style.borderTopColor) ?? { hex: "#C8B9A7", alpha: 1 },
      border: null,
      order: nextOrder(),
      zIndex
    });
    return;
  }

  if (!isRoot && hasBackgroundImage(style.backgroundImage)) {
    const backgroundImage = buildBackgroundImageElement(element, bounds, style.backgroundImage, collector.warnings, nextOrder(), zIndex);
    if (backgroundImage) {
      collector.push(backgroundImage);
    }
  }

  if (!isRoot && hasVisibleBox(style)) {
    const shapeElement = buildShapeElement(rect, bounds, style, nextOrder(), zIndex);
    if (shapeElement) {
      collector.push(shapeElement);
    }
  }

  if (hasInlineMediaText(element, childElements)) {
    const textOffsetLeft = calculateInlineMediaOffset(element, childElements.filter(isInlineMediaElement));
    const textElement = buildTextElement(element, bounds, style, nextOrder(), zIndex, textOffsetLeft);
    if (textElement) {
      collector.push(textElement);
    }

    childElements.filter(isInlineMediaElement).forEach((child) => {
      traverseElement(child, bounds, collector, nextOrder, false);
    });
    return;
  }

  if (isLeafTextElement(element)) {
    const textElement = buildTextElement(element, bounds, style, nextOrder(), zIndex);
    if (textElement) {
      collector.push(textElement);
    }
    return;
  }

  if (!childElements.length) {
    const textElement = buildTextElement(element, bounds, style, nextOrder(), zIndex);
    if (textElement) {
      collector.push(textElement);
    }
    return;
  }

  for (const child of childElements) {
    traverseElement(child, bounds, collector, nextOrder, false);
  }
}

function buildMediaElement(
  element: HTMLElement,
  bounds: Bounds,
  warnings: string[],
  order: number,
  zIndex: number
): ImageElement | null {
  const rect = getAbsoluteRect(element);
  const src = resolveMediaSource(element, warnings);

  if (!src) {
    return null;
  }

  return {
    type: "image",
    x: rect.left - bounds.left,
    y: rect.top - bounds.top,
    w: rect.width,
    h: rect.height,
    src,
    alt: element.getAttribute("alt") ?? undefined,
    order,
    zIndex
  };
}

function buildBackgroundImageElement(
  element: HTMLElement,
  bounds: Bounds,
  backgroundImage: string,
  warnings: string[],
  order: number,
  zIndex: number
): ImageElement | null {
  const firstUrl = extractFirstCssUrl(backgroundImage);
  if (!firstUrl) {
    return null;
  }

  const resolvedUrl = resolveUrlAgainstDocument(firstUrl, element.ownerDocument);
  if (!resolvedUrl) {
    warnings.push("Uma imagem de fundo relativa nao pode ser incorporada automaticamente sem estar inline no HTML.");
    return null;
  }

  const rect = getAbsoluteRect(element);
  return {
    type: "image",
    x: rect.left - bounds.left,
    y: rect.top - bounds.top,
    w: rect.width,
    h: rect.height,
    src: resolvedUrl,
    order,
    zIndex
  };
}

function buildShapeElement(
  rect: DOMRect,
  bounds: Bounds,
  style: CSSStyleDeclaration,
  order: number,
  zIndex: number
): ShapeElement | null {
  const fill = toColorToken(style.backgroundColor);
  const border = extractBorder(style);
  const radius = Math.max(
    parsePx(style.borderTopLeftRadius),
    parsePx(style.borderTopRightRadius),
    parsePx(style.borderBottomLeftRadius),
    parsePx(style.borderBottomRightRadius)
  );

  if (!fill && !border) {
    return null;
  }

  return {
    type: "shape",
    kind: radius > 2 ? "roundRect" : "rect",
    x: rect.left - bounds.left,
    y: rect.top - bounds.top,
    w: rect.width,
    h: rect.height,
    radiusPx: radius || undefined,
    fill,
    border,
    order,
    zIndex
  };
}

function buildTextElement(
  element: HTMLElement,
  bounds: Bounds,
  style: CSSStyleDeclaration,
  order: number,
  zIndex: number,
  textOffsetLeft = 0
): TextElement | null {
  const preserveWhitespace = element.tagName === "PRE";
  const text = normalizeText(element.innerText || element.textContent || "", preserveWhitespace);

  if (!text) {
    return null;
  }

  const rect = getAbsoluteRect(element);
  const padding = {
    top: parsePx(style.paddingTop),
    right: parsePx(style.paddingRight),
    bottom: parsePx(style.paddingBottom),
    left: parsePx(style.paddingLeft)
  };

  const contentX = rect.left - bounds.left + padding.left + textOffsetLeft;
  const contentY = rect.top - bounds.top + padding.top;
  const contentW = Math.max(rect.width - padding.left - padding.right - textOffsetLeft, 1);
  const contentH = Math.max(rect.height - padding.top - padding.bottom, 1);

  const fontSizePx = parsePx(style.fontSize) || 16;
  const weight = parseFontWeight(style.fontWeight);
  const isBold = weight >= 600;
  const isItalic = style.fontStyle === "italic";
  const isUnderlined = style.textDecorationLine.includes("underline");

  return {
    type: "text",
    x: contentX,
    y: contentY,
    w: contentW,
    h: contentH,
    text,
    runs: [
      {
        text,
        color: toColorToken(style.color),
        fontFamily: normalizeFontFamily(style.fontFamily),
        fontSizePx,
        bold: isBold,
        italic: isItalic,
        underline: isUnderlined
      }
    ],
    fontFamily: normalizeFontFamily(style.fontFamily),
    fontSizePx,
    color: toColorToken(style.color),
    align: mapHorizontalAlign(style.textAlign),
    verticalAlign: mapVerticalAlign(style.verticalAlign),
    bold: isBold,
    italic: isItalic,
    underline: isUnderlined,
    order,
    zIndex
  };
}

function hasInlineMediaText(element: HTMLElement, childElements: HTMLElement[]): boolean {
  if (!hasTextContent(element) || !childElements.length) {
    return false;
  }

  const mediaChildren = childElements.filter(isInlineMediaElement);
  if (!mediaChildren.length) {
    return false;
  }

  return childElements.every((child) => INLINE_TAGS.has(child.tagName) || isInlineMediaElement(child));
}

function isInlineMediaElement(element: HTMLElement): boolean {
  return ["IMG", "SVG", "CANVAS"].includes(element.tagName);
}

function calculateInlineMediaOffset(element: HTMLElement, mediaChildren: HTMLElement[]): number {
  if (!mediaChildren.length) {
    return 0;
  }

  const parentRect = getAbsoluteRect(element);
  const gap = parsePx(getComputedStyle(element).columnGap) || parsePx(getComputedStyle(element).gap);
  const farthestRight = Math.max(...mediaChildren.map((child) => getAbsoluteRect(child).right));
  return Math.max(0, farthestRight - parentRect.left + gap);
}

function buildListElements(
  list: HTMLUListElement | HTMLOListElement,
  bounds: Bounds,
  collector: ElementCollector,
  nextOrder: () => number,
  zIndex: number,
  depth = 0
): void {
  const items = Array.from(list.children).filter((child): child is HTMLLIElement => child instanceof HTMLLIElement);
  const ordered = list.tagName === "OL";
  let currentNumber = parseInt(list.getAttribute("start") ?? "1", 10);

  items.forEach((item) => {
    const itemStyle = getComputedStyle(item);
    const cloned = item.cloneNode(true) as HTMLLIElement;
    cloned.querySelectorAll("ul, ol").forEach((node) => node.remove());
    const rawText = normalizeText(cloned.innerText || cloned.textContent || "", false);

    if (rawText) {
      const rect = getAbsoluteRect(item);
      const prefix = ordered ? `${Number.isNaN(currentNumber) ? 1 : currentNumber}. ` : `${"  ".repeat(depth)}• `;

      collector.push({
        type: "text",
        x: rect.left - bounds.left,
        y: rect.top - bounds.top,
        w: rect.width,
        h: rect.height,
        text: `${prefix}${rawText}`,
        runs: [
          {
            text: `${prefix}${rawText}`,
            color: toColorToken(itemStyle.color),
            fontFamily: normalizeFontFamily(itemStyle.fontFamily),
            fontSizePx: parsePx(itemStyle.fontSize) || 16,
            bold: parseFontWeight(itemStyle.fontWeight) >= 600,
            italic: itemStyle.fontStyle === "italic",
            underline: itemStyle.textDecorationLine.includes("underline")
          }
        ],
        fontFamily: normalizeFontFamily(itemStyle.fontFamily),
        fontSizePx: parsePx(itemStyle.fontSize) || 16,
        color: toColorToken(itemStyle.color),
        align: mapHorizontalAlign(itemStyle.textAlign),
        verticalAlign: "top",
        bold: parseFontWeight(itemStyle.fontWeight) >= 600,
        italic: itemStyle.fontStyle === "italic",
        underline: itemStyle.textDecorationLine.includes("underline"),
        order: nextOrder(),
        zIndex
      });
    }

    if (ordered) {
      currentNumber += 1;
    }

    const nestedLists = Array.from(item.children).filter(
      (child): child is HTMLUListElement | HTMLOListElement => child instanceof HTMLUListElement || child instanceof HTMLOListElement
    );

    nestedLists.forEach((nestedList) => buildListElements(nestedList, bounds, collector, nextOrder, zIndex + 1, depth + 1));
  });
}

function buildTableElement(table: HTMLTableElement, bounds: Bounds, order: number, zIndex: number): TableElement | null {
  if (!table.rows.length) {
    return null;
  }

  const rect = getAbsoluteRect(table);
  const rows = Array.from(table.rows).map((row) =>
    Array.from(row.cells).map((cell) => buildTableCell(cell))
  );

  const firstRow = table.rows[0];
  const colWidths = Array.from(firstRow.cells).map((cell) => getAbsoluteRect(cell).width);
  const rowHeights = Array.from(table.rows).map((row) => getAbsoluteRect(row).height);

  return {
    type: "table",
    x: rect.left - bounds.left,
    y: rect.top - bounds.top,
    w: rect.width,
    h: rect.height,
    colWidths,
    rowHeights,
    rows,
    border: extractBorder(getComputedStyle(table)),
    fill: toColorToken(getComputedStyle(table).backgroundColor),
    order,
    zIndex
  };
}

function buildTableCell(cell: HTMLTableCellElement): TableCellManifest {
  const style = getComputedStyle(cell);
  return {
    text: normalizeText(cell.innerText || cell.textContent || "", false),
    background: toColorToken(style.backgroundColor),
    color: toColorToken(style.color),
    fontFamily: normalizeFontFamily(style.fontFamily),
    fontSizePx: parsePx(style.fontSize) || 14,
    align: mapHorizontalAlign(style.textAlign),
    bold: parseFontWeight(style.fontWeight) >= 600
  };
}

function resolveMediaSource(element: HTMLElement, warnings: string[]): string | null {
  if (element instanceof HTMLImageElement) {
    const rawSource = element.currentSrc || element.getAttribute("src") || "";
    const resolvedUrl = resolveUrlAgainstDocument(rawSource, element.ownerDocument);

    if (!resolvedUrl) {
      warnings.push("Uma imagem relativa nao pode ser enviada para o PPTX sem estar embutida em data URL ou acessivel por URL publica.");
      return null;
    }

    return resolvedUrl;
  }

  if (element instanceof SVGElement) {
    const markup = new XMLSerializer().serializeToString(element);
    return `data:image/svg+xml;base64,${toBase64Utf8(markup)}`;
  }

  if (element instanceof HTMLCanvasElement) {
    try {
      return element.toDataURL("image/png");
    } catch {
      warnings.push("Um elemento canvas nao pode ser exportado porque o navegador bloqueou a leitura do bitmap.");
      return null;
    }
  }

  return null;
}

function resolveUrlAgainstDocument(rawSource: string, doc: Document): string | null {
  if (!rawSource) {
    return null;
  }

  if (rawSource.startsWith("data:")) {
    return rawSource;
  }

  try {
    const resolved = new URL(rawSource, doc.baseURI).toString();
    if (resolved.startsWith("about:")) {
      return null;
    }
    return resolved;
  } catch {
    return null;
  }
}

function extractFirstCssUrl(backgroundImage: string): string | null {
  const match = backgroundImage.match(/url\((['"]?)(.*?)\1\)/i);
  return match?.[2] ?? null;
}

function hasBackgroundImage(backgroundImage: string): boolean {
  return backgroundImage !== "none" && backgroundImage.includes("url(");
}

function hasVisibleBox(style: CSSStyleDeclaration): boolean {
  return Boolean(toColorToken(style.backgroundColor) || extractBorder(style));
}

function extractBorder(style: CSSStyleDeclaration): BorderToken | null {
  const borderWidth = parsePx(style.borderTopWidth);
  const borderStyle = style.borderTopStyle;
  const borderColor = toColorToken(style.borderTopColor);

  if (!borderWidth || borderStyle === "none" || !borderColor) {
    return null;
  }

  return {
    color: borderColor,
    widthPx: borderWidth
  };
}

function isLeafTextElement(element: HTMLElement): boolean {
  const tag = element.tagName;

  if (["UL", "OL", "TABLE", "IMG", "SVG", "CANVAS"].includes(tag)) {
    return false;
  }

  const children = Array.from(element.children).filter((child): child is HTMLElement => child instanceof HTMLElement);
  if (!children.length) {
    return hasTextContent(element);
  }

  if (!hasTextContent(element)) {
    return false;
  }

  return children.every((child) => INLINE_TAGS.has(child.tagName));
}

function hasTextContent(element: HTMLElement): boolean {
  return normalizeText(element.innerText || element.textContent || "", element.tagName === "PRE").length > 0;
}

function isVisibleElement(element: HTMLElement): boolean {
  const style = getComputedStyle(element);
  if (style.display === "none" || style.visibility === "hidden" || Number(style.opacity) === 0) {
    return false;
  }

  const rect = element.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

function intersectsBounds(rect: DOMRect, bounds: Bounds): boolean {
  const boundsBottom = bounds.top + bounds.height;
  const boundsRight = bounds.left + bounds.width;

  return rect.left < boundsRight && rect.right > bounds.left && rect.top < boundsBottom && rect.bottom > bounds.top;
}

function getAbsoluteRect(element: Element): DOMRect {
  const rect = element.getBoundingClientRect();
  return new DOMRect(rect.left, rect.top, rect.width, rect.height);
}

function parsePx(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseFontWeight(value: string): number {
  if (value === "bold") {
    return 700;
  }

  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : 400;
}

function parseZIndex(value: string): number {
  if (value === "auto") {
    return 0;
  }

  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeFontFamily(fontFamily: string): string | null {
  const firstFamily = fontFamily
    .split(",")
    .map((item) => item.trim().replace(/^['"]|['"]$/g, ""))
    .find(Boolean);

  return firstFamily || null;
}

function normalizeText(value: string, preserveWhitespace: boolean): string {
  const withoutCarriageReturn = value.replace(/\r/g, "").replace(/\u00a0/g, " ");

  if (preserveWhitespace) {
    return withoutCarriageReturn.trimEnd();
  }

  return withoutCarriageReturn
    .split("\n")
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter((line, index, lines) => line.length > 0 || (index > 0 && index < lines.length - 1))
    .join("\n")
    .trim();
}

function mapHorizontalAlign(value: string): HorizontalAlign {
  if (value === "center") {
    return "center";
  }

  if (value === "right" || value === "end") {
    return "right";
  }

  if (value === "justify") {
    return "justify";
  }

  return "left";
}

function mapVerticalAlign(value: string): VerticalAlign {
  if (value === "middle" || value === "center") {
    return "middle";
  }

  if (value === "bottom") {
    return "bottom";
  }

  return "top";
}

function toColorToken(value: string | null): ColorToken | null {
  if (!value || value === "transparent") {
    return null;
  }

  const rgbaMatch = value.match(
    /rgba?\(\s*([0-9.]+)\s*,\s*([0-9.]+)\s*,\s*([0-9.]+)(?:\s*,\s*([0-9.]+))?\s*\)/i
  );

  if (!rgbaMatch) {
    if (value.startsWith("#")) {
      return {
        hex: value.length === 4
          ? `#${value[1]}${value[1]}${value[2]}${value[2]}${value[3]}${value[3]}`
          : value.toUpperCase(),
        alpha: 1
      };
    }

    return null;
  }

  const red = clampColorChannel(Number(rgbaMatch[1]));
  const green = clampColorChannel(Number(rgbaMatch[2]));
  const blue = clampColorChannel(Number(rgbaMatch[3]));
  const alpha = rgbaMatch[4] ? Math.max(0, Math.min(1, Number(rgbaMatch[4]))) : 1;

  if (alpha === 0) {
    return null;
  }

  return {
    hex: `#${toHex(red)}${toHex(green)}${toHex(blue)}`,
    alpha
  };
}

function clampColorChannel(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function toHex(value: number): string {
  return value.toString(16).padStart(2, "0").toUpperCase();
}

function dedupeMessages(messages: string[]): string[] {
  return Array.from(new Set(messages.filter(Boolean)));
}

function extractTitleFromDocument(doc: Document): string {
  const fromTitle = doc.title?.trim();
  if (fromTitle) {
    return fromTitle;
  }

  const firstHeading = doc.body.querySelector("h1, h2");
  return firstHeading?.textContent?.trim() || "Apresentacao";
}

function inferSlideTitle(root: HTMLElement | undefined, index: number): string {
  if (!root) {
    return `Slide ${index}`;
  }

  const heading = root.querySelector("h1, h2, h3");
  if (heading?.textContent?.trim()) {
    return heading.textContent.trim();
  }

  const ariaLabel = root.getAttribute("aria-label")?.trim();
  if (ariaLabel) {
    return ariaLabel;
  }

  return `Slide ${index}`;
}

function toOutputFileName(sourceName: string): string {
  const sanitized = sourceName.replace(/\.[^.]+$/, "").trim() || "apresentacao";
  return `${sanitized}.pptx`;
}

function toBase64Utf8(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";

  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });

  return btoa(binary);
}
