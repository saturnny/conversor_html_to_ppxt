export const EXPORT_FRAME_API_KEY = "__HTML_TO_PPTX_EXPORT__" as const;

const EXPLICIT_SLIDE_SELECTOR =
  ".presentation-container > .slide, body > .slide, .slide, [data-slide], [data-ppt-slide], .ppt-slide, .page-slide";

const HELPER_ATTRIBUTE = "data-html-to-pptx-helper";
const BUNDLE_ENDPOINT = "/api/dom-to-pptx-bundle";

const FRAME_SUPPORT_STYLES = `
  [${HELPER_ATTRIBUTE}="stage"] {
    position: fixed !important;
    left: -100000px !important;
    top: 0 !important;
    width: 1920px !important;
    display: flex !important;
    flex-direction: column !important;
    gap: 32px !important;
    pointer-events: none !important;
    opacity: 0 !important;
    z-index: -1 !important;
  }

  [${HELPER_ATTRIBUTE}="slide"] {
    position: relative !important;
    margin: 0 !important;
    pointer-events: none !important;
  }
`;

export type EditableFontDefinition = {
  name: string;
  url: string;
};

export type EditableListConfig = {
  color?: string;
  spacing?: {
    before?: number;
    after?: number;
  };
};

export type PreparedEditableExport = {
  srcDoc: string;
  downloadName: string;
  warnings: string[];
};

export type FrameInspectionResult = {
  slideCount: number;
  warnings: string[];
};

export type FrameExportResult = {
  blob: Blob;
  fileName: string;
  slideCount: number;
  warnings: string[];
};

export type ExportFrameApi = {
  inspect: () => Promise<FrameInspectionResult>;
  exportSlides: (options?: {
    fileName?: string;
    autoEmbedFonts?: boolean;
    fonts?: EditableFontDefinition[];
    svgAsVector?: boolean;
    listConfig?: EditableListConfig;
  }) => Promise<FrameExportResult>;
};

export type ExportFrameWindow = Window & {
  [EXPORT_FRAME_API_KEY]?: ExportFrameApi;
};

export async function decodeHtmlFile(file: File): Promise<string> {
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

export function prepareEditableExportDocument(
  html: string,
  sourceName: string
): PreparedEditableExport {
  const warnings: string[] = [];
  const parser = new DOMParser();
  const sourceDocument = parser.parseFromString(extractRenderableHtml(html), "text/html");
  const removedScripts = sanitizeDocument(sourceDocument);

  if (removedScripts > 0) {
    warnings.push(
      "Scripts foram removidos do HTML por segurança. Se o layout depender de JavaScript para montar o slide, a exportação pode variar."
    );
  }

  if (hasRelativeAssets(sourceDocument)) {
    warnings.push(
      "Foram detectados assets com caminho relativo. Imagens, fontes ou CSS locais precisam estar em URL pública ou embutidos no HTML."
    );
  }

  if (hasGoogleFontsWithoutCrossOrigin(sourceDocument)) {
    warnings.push(
      "Foram detectados links do Google Fonts sem crossorigin. O PowerPoint pode trocar a fonte caso a incorporação automática falhe."
    );
  }

  sourceDocument.documentElement.setAttribute("lang", sourceDocument.documentElement.lang || "pt-BR");
  sourceDocument.body.setAttribute("data-html-to-pptx-source", sourceName);
  ensureCrossOriginFontLinks(sourceDocument);

  const supportStyle = sourceDocument.createElement("style");
  supportStyle.setAttribute(HELPER_ATTRIBUTE, "support-style");
  supportStyle.textContent = FRAME_SUPPORT_STYLES;
  sourceDocument.head.appendChild(supportStyle);

  const bundleScript = sourceDocument.createElement("script");
  bundleScript.setAttribute(HELPER_ATTRIBUTE, "bundle");
  bundleScript.src = BUNDLE_ENDPOINT;
  sourceDocument.body.appendChild(bundleScript);

  const helperScript = sourceDocument.createElement("script");
  helperScript.setAttribute(HELPER_ATTRIBUTE, "bridge");
  helperScript.textContent = buildFrameBridgeScript(toOutputFileName(sourceName));
  sourceDocument.body.appendChild(helperScript);

  const markup = `<!DOCTYPE html>\n${sourceDocument.documentElement.outerHTML}`;

  return {
    srcDoc: markup,
    downloadName: toOutputFileName(sourceName),
    warnings
  };
}

function buildFrameBridgeScript(defaultFileName: string): string {
  return `
    (function () {
      const apiKey = ${JSON.stringify(EXPORT_FRAME_API_KEY)};
      const explicitSelector = ${JSON.stringify(EXPLICIT_SLIDE_SELECTOR)};
      const helperAttribute = ${JSON.stringify(HELPER_ATTRIBUTE)};
      const stageId = "__html-to-pptx-export-stage";
      const navigationSelector = ".nav-btn, .nav-arrow, [data-nav], [data-slide-nav], [aria-label*='next' i], [aria-label*='previous' i], [aria-label*='anterior' i], [aria-label*='proximo' i]";
      const indicatorSelector = ".page-indicator, .slide-counter, [data-slide-indicator], [data-page-indicator]";
      const progressSelector = ".footer-progress-bar, [data-slide-progress], .progress-bar__fill";
      const fallbackFileName = ${JSON.stringify(defaultFileName)};
      const assetProxyBase = "/api/asset-proxy?url=";

      function waitForImage(image) {
        return new Promise((resolve) => {
          if (image.complete) {
            resolve();
            return;
          }

          const done = () => resolve();
          image.addEventListener("load", done, { once: true });
          image.addEventListener("error", done, { once: true });
          window.setTimeout(done, 2000);
        });
      }

      async function waitForDocumentReady() {
        if (document.fonts && document.fonts.ready) {
          try {
            await document.fonts.ready;
          } catch {}
        }

        const images = Array.from(document.images || []);
        await Promise.all(images.map(waitForImage));
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      }

      function getPrimaryFontFamily(value) {
        if (!value) {
          return "";
        }

        return String(value).split(",")[0].replace(/['"]/g, "").trim();
      }

      function collectUsedFontFamilies(targets) {
        const families = new Set();

        targets.forEach((target) => {
          const nodes = [target, ...Array.from(target.querySelectorAll("*"))];
          nodes.forEach((node) => {
            if (!(node instanceof HTMLElement)) {
              return;
            }

            const family = getPrimaryFontFamily(window.getComputedStyle(node).fontFamily);
            if (family) {
              families.add(family);
            }
          });
        });

        return families;
      }

      function extractFontUrl(srcText) {
        if (!srcText) {
          return null;
        }

        const matches = srcText.match(/url\\((['"]?)(.*?)\\1\\)/g);
        if (!matches) {
          return null;
        }

        let fallbackUrl = null;

        for (const match of matches) {
          const url = match.replace(/url\\((['"]?)(.*?)\\1\\)/, "$2");
          if (!url || url.startsWith("data:")) {
            continue;
          }

          if (url.includes(".ttf") || url.includes(".otf") || url.includes(".woff2") || url.includes(".woff")) {
            return url;
          }

          if (!fallbackUrl) {
            fallbackUrl = url;
          }
        }

        return fallbackUrl;
      }

      function collectExplicitFontDefinitions(targets) {
        const usedFamilies = collectUsedFontFamilies(targets);
        const collected = [];
        const processed = new Set();

        Array.from(document.styleSheets || []).forEach((sheet) => {
          try {
            const rules = sheet.cssRules || sheet.rules;
            if (!rules) {
              return;
            }

            Array.from(rules).forEach((rule) => {
              if (!(rule instanceof CSSFontFaceRule)) {
                return;
              }

              const family = getPrimaryFontFamily(rule.style.getPropertyValue("font-family"));
              if (!family || !usedFamilies.has(family)) {
                return;
              }

              const url = extractFontUrl(rule.style.getPropertyValue("src"));
              if (!url || processed.has(url)) {
                return;
              }

              processed.add(url);
              collected.push({ name: family, url: url });
            });
          } catch {}
        });

        return collected;
      }

      function dedupeFonts(fonts) {
        const seen = new Set();

        return fonts.filter((font) => {
          if (!font || !font.name || !font.url) {
            return false;
          }

          const key = String(font.name).trim().toLowerCase() + "::" + String(font.url).trim();
          if (seen.has(key)) {
            return false;
          }

          seen.add(key);
          return true;
        });
      }

      function toAbsoluteUrl(value) {
        if (!value) {
          return null;
        }

        try {
          return new URL(value, window.location.href).toString();
        } catch {
          return null;
        }
      }

      function proxifyUrl(value) {
        const absoluteUrl = toAbsoluteUrl(value);
        if (!absoluteUrl) {
          return value;
        }

        try {
          const parsed = new URL(absoluteUrl);
          const isHttp = parsed.protocol === "http:" || parsed.protocol === "https:";
          if (!isHttp || parsed.origin === window.location.origin) {
            return absoluteUrl;
          }

          return assetProxyBase + encodeURIComponent(absoluteUrl);
        } catch {
          return value;
        }
      }

      function splitCssLayers(value) {
        const input = String(value || "").trim();
        if (!input) {
          return [];
        }

        const result = [];
        let current = "";
        let depth = 0;
        let quote = "";

        for (let index = 0; index < input.length; index += 1) {
          const character = input[index];

          if (quote) {
            current += character;
            if (character === quote && input[index - 1] !== "\\\\") {
              quote = "";
            }
            continue;
          }

          if (character === "'" || character === '"') {
            quote = character;
            current += character;
            continue;
          }

          if (character === "(") {
            depth += 1;
            current += character;
            continue;
          }

          if (character === ")") {
            depth = Math.max(0, depth - 1);
            current += character;
            continue;
          }

          if (character === "," && depth === 0) {
            result.push(current.trim());
            current = "";
            continue;
          }

          current += character;
        }

        if (current.trim()) {
          result.push(current.trim());
        }

        return result;
      }

      function extractBackgroundUrl(layer) {
        const match = String(layer || "").match(/url\\((['"]?)(.*?)\\1\\)/i);
        return match?.[2] || null;
      }

      function rewriteLayerUrl(layer) {
        const assetUrl = extractBackgroundUrl(layer);
        if (!assetUrl) {
          return layer;
        }

        return String(layer).replace(assetUrl, proxifyUrl(assetUrl));
      }

      function rewriteImageSources(root) {
        const candidates = [];

        if (root instanceof HTMLElement && root.matches("img[src]")) {
          candidates.push(root);
        }

        root.querySelectorAll("img[src]").forEach((image) => {
          if (image instanceof HTMLImageElement) {
            candidates.push(image);
          }
        });

        candidates.forEach((image) => {
          const source = image.getAttribute("src");
          if (!source) {
            return;
          }

          const proxied = proxifyUrl(source);
          if (proxied && proxied !== source) {
            image.setAttribute("src", proxied);
          }
        });
      }

      function normalizeBackgroundAssets(root) {
        const elements = [root, ...Array.from(root.querySelectorAll("*"))];

        elements.forEach((element) => {
          if (!(element instanceof HTMLElement)) {
            return;
          }

          const style = window.getComputedStyle(element);
          const layers = splitCssLayers(style.backgroundImage).filter((layer) => {
            return layer && layer !== "none";
          });

          if (!layers.some((layer) => layer.includes("url("))) {
            return;
          }

          const positions = splitCssLayers(style.backgroundPosition);
          const sizes = splitCssLayers(style.backgroundSize);
          const repeats = splitCssLayers(style.backgroundRepeat);

          if (style.position === "static") {
            element.style.setProperty("position", "relative", "important");
          }

          element.style.setProperty("isolation", "isolate", "important");
          element.style.setProperty("background-image", "none", "important");

          layers
            .map((layer, index) => ({
              layer,
              index
            }))
            .reverse()
            .forEach(({ layer, index }, reverseIndex) => {
              const backgroundLayer = document.createElement("div");
              backgroundLayer.setAttribute(helperAttribute, "background-layer");
              backgroundLayer.style.position = "absolute";
              backgroundLayer.style.inset = "0";
              backgroundLayer.style.pointerEvents = "none";
              backgroundLayer.style.zIndex = String(-(reverseIndex + 1));
              backgroundLayer.style.borderRadius = style.borderRadius;
              backgroundLayer.style.backgroundImage = rewriteLayerUrl(layer);
              backgroundLayer.style.backgroundPosition = positions[index] || positions[positions.length - 1] || "center";
              backgroundLayer.style.backgroundSize = sizes[index] || sizes[sizes.length - 1] || "auto";
              backgroundLayer.style.backgroundRepeat = repeats[index] || repeats[repeats.length - 1] || "repeat";
              backgroundLayer.style.backgroundColor = "transparent";

              element.insertBefore(backgroundLayer, element.firstChild);
            });
        });
      }

      function filterNested(elements) {
        return elements.filter((element, index) => {
          return !elements.some((candidate, candidateIndex) => {
            return candidateIndex !== index && candidate.contains(element);
          });
        });
      }

      function defaultDisplayFor(tagName) {
        switch (String(tagName || "").toLowerCase()) {
          case "span":
            return "inline";
          case "li":
            return "list-item";
          case "table":
            return "table";
          case "thead":
            return "table-header-group";
          case "tbody":
            return "table-row-group";
          case "tr":
            return "table-row";
          case "td":
          case "th":
            return "table-cell";
          case "svg":
            return "inline-block";
          default:
            return "block";
        }
      }

      function getExplicitSlides() {
        const candidates = Array.from(document.querySelectorAll(explicitSelector)).filter((element) => {
          return element instanceof HTMLElement;
        });

        return filterNested(candidates);
      }

      function hideNavigation(root) {
        const scope = root && typeof root.querySelectorAll === "function" ? root : document;

        scope.querySelectorAll(navigationSelector).forEach((element) => {
          if (!(element instanceof HTMLElement)) {
            return;
          }

          element.style.setProperty("display", "none", "important");
          element.style.setProperty("visibility", "hidden", "important");
        });
      }

      function updateIndicators(root, index, total) {
        root.querySelectorAll(indicatorSelector).forEach((element) => {
          if (element instanceof HTMLElement) {
            element.textContent = String(index + 1) + " / " + String(total);
          }
        });

        root.querySelectorAll(progressSelector).forEach((element) => {
          if (element instanceof HTMLElement) {
            element.style.setProperty("width", String(((index + 1) / total) * 100) + "%", "important");
          }
        });
      }

      function computeSlideDisplay(slide) {
        slide.style.removeProperty("display");
        const computedDisplay = window.getComputedStyle(slide).display;
        if (computedDisplay && computedDisplay !== "none") {
          return computedDisplay;
        }

        return defaultDisplayFor(slide.tagName);
      }

      function setActiveSlide(slides, targetIndex, total) {
        slides.forEach((slide, slideIndex) => {
          const isActive = slideIndex === targetIndex;

          slide.classList.toggle("active", isActive);
          slide.classList.toggle("is-active", isActive);
          slide.classList.toggle("current", isActive);
          slide.classList.add("print-visible");
          slide.removeAttribute("hidden");

          if (isActive) {
            const display = computeSlideDisplay(slide);
            slide.style.setProperty("display", display, "important");
            slide.style.setProperty("visibility", "visible", "important");
            slide.style.setProperty("opacity", "1", "important");
            slide.style.setProperty("z-index", "999", "important");
          } else {
            slide.style.setProperty("display", "none", "important");
            slide.style.setProperty("visibility", "hidden", "important");
            slide.style.setProperty("opacity", "0", "important");
            slide.style.setProperty("z-index", "0", "important");
          }

          slide.style.setProperty("pointer-events", "none", "important");
          updateIndicators(slide, targetIndex, total);
        });

        hideNavigation(document);
      }

      function getSlideWrapperSize(slide) {
        const rect = slide.getBoundingClientRect();
        const width = Math.max(Math.round(rect.width), 1);
        const height = Math.max(Math.round(rect.height), 1);

        return { width, height };
      }

      function buildSlideWrapper(slide, index, total) {
        const display = computeSlideDisplay(slide);
        const size = getSlideWrapperSize(slide);
        const wrapper = document.createElement("section");
        wrapper.setAttribute(helperAttribute, "slide");
        wrapper.style.position = "relative";
        wrapper.style.overflow = "hidden";
        wrapper.style.width = String(size.width) + "px";
        wrapper.style.height = String(size.height) + "px";

        const clone = slide.cloneNode(true);
        clone.classList.add("active", "is-active", "current", "print-visible");
        clone.removeAttribute("hidden");
        clone.style.setProperty("display", display, "important");
        clone.style.setProperty("visibility", "visible", "important");
        clone.style.setProperty("opacity", "1", "important");
        clone.style.setProperty("position", "relative", "important");
        clone.style.setProperty("left", "0", "important");
        clone.style.setProperty("top", "0", "important");
        clone.style.setProperty("inset", "auto", "important");
        clone.style.setProperty("margin", "0", "important");
        clone.style.setProperty("width", String(size.width) + "px", "important");
        clone.style.setProperty("height", String(size.height) + "px", "important");
        clone.style.setProperty("pointer-events", "none", "important");

        rewriteImageSources(clone);
        normalizeBackgroundAssets(clone);
        hideNavigation(clone);
        updateIndicators(clone, index, total);
        wrapper.appendChild(clone);

        return wrapper;
      }

      function buildFallbackSlide() {
        const bodyRect = document.body.getBoundingClientRect();
        const wrapper = document.createElement("section");
        wrapper.setAttribute(helperAttribute, "slide");
        wrapper.style.position = "relative";
        wrapper.style.overflow = "hidden";
        wrapper.style.width = String(Math.max(Math.round(bodyRect.width), document.documentElement.clientWidth || 1280, 1280)) + "px";
        wrapper.style.height = String(Math.max(Math.round(bodyRect.height), document.documentElement.clientHeight || 720, 720)) + "px";

        const container = document.createElement("div");
        container.style.position = "relative";
        container.style.width = "100%";
        container.style.height = "100%";

        Array.from(document.body.children).forEach((child) => {
          if (!(child instanceof HTMLElement)) {
            return;
          }

          if (child.hasAttribute(helperAttribute) || child.id === stageId) {
            return;
          }

          container.appendChild(child.cloneNode(true));
        });

        rewriteImageSources(container);
        normalizeBackgroundAssets(container);
        hideNavigation(container);
        wrapper.appendChild(container);

        return wrapper;
      }

      function ensureExportStage() {
        const existing = document.getElementById(stageId);
        if (existing) {
          existing.remove();
        }

        const stage = document.createElement("div");
        stage.id = stageId;
        stage.setAttribute(helperAttribute, "stage");
        document.body.appendChild(stage);
        return stage;
      }

      function buildSlidesForExport() {
        const stage = ensureExportStage();
        const warnings = [];
        const slides = getExplicitSlides();

        if (slides.length > 1) {
          const targets = slides.map((slide, index) => {
            setActiveSlide(slides, index, slides.length);
            const wrapper = buildSlideWrapper(slide, index, slides.length);
            stage.appendChild(wrapper);
            return wrapper;
          });

          return {
            targets,
            slideCount: targets.length,
            warnings
          };
        }

        if (slides.length === 1) {
          setActiveSlide(slides, 0, 1);
          const wrapper = buildSlideWrapper(slides[0], 0, 1);
          stage.appendChild(wrapper);

          return {
            targets: [wrapper],
            slideCount: 1,
            warnings
          };
        }

        warnings.push("Nenhum seletor explicito de slide foi encontrado. O corpo do HTML sera exportado como um unico slide editavel.");

        const fallback = buildFallbackSlide();
        stage.appendChild(fallback);

        return {
          targets: [fallback],
          slideCount: 1,
          warnings
        };
      }

      async function inspect() {
        await waitForDocumentReady();
        const prepared = buildSlidesForExport();
        return {
          slideCount: prepared.slideCount,
          warnings: prepared.warnings
        };
      }

      async function exportSlides(options) {
        await waitForDocumentReady();

        if (!window.domToPptx || typeof window.domToPptx.exportToPptx !== "function") {
          throw new Error("O bundle do dom-to-pptx nao foi carregado no sandbox de exportacao.");
        }

        const prepared = buildSlidesForExport();
        const exportOptions = Object.assign(
          {
            fileName: fallbackFileName,
            autoEmbedFonts: true,
            fonts: [],
            svgAsVector: true,
            skipDownload: true
          },
          options || {}
        );

        const inferredFonts = collectExplicitFontDefinitions(prepared.targets);
        exportOptions.fonts = dedupeFonts([...(exportOptions.fonts || []), ...inferredFonts]);

        const blob = await window.domToPptx.exportToPptx(prepared.targets, exportOptions);

        return {
          blob,
          fileName: exportOptions.fileName || fallbackFileName,
          slideCount: prepared.slideCount,
          warnings: prepared.warnings
        };
      }

      window[apiKey] = {
        inspect,
        exportSlides
      };
    })();
  `;
}

function ensureCrossOriginFontLinks(documentNode: Document) {
  Array.from(documentNode.querySelectorAll('link[href*="fonts.googleapis.com"], link[href*="fonts.gstatic.com"]')).forEach(
    (element) => {
      if (element instanceof HTMLLinkElement && !element.crossOrigin) {
        element.crossOrigin = "anonymous";
      }
    }
  );
}

function sanitizeDocument(documentNode: Document): number {
  const scripts = Array.from(documentNode.querySelectorAll("script, noscript"));
  scripts.forEach((node) => node.remove());

  documentNode
    .querySelectorAll('meta[http-equiv="Content-Security-Policy"], meta[http-equiv="refresh"], base')
    .forEach((node) => node.remove());

  documentNode.querySelectorAll("*").forEach((element) => {
    Array.from(element.attributes).forEach((attribute) => {
      if (/^on/i.test(attribute.name)) {
        element.removeAttribute(attribute.name);
        return;
      }

      if ((attribute.name === "href" || attribute.name === "src") && /^javascript:/i.test(attribute.value)) {
        element.removeAttribute(attribute.name);
      }
    });
  });

  return scripts.length;
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

function toOutputFileName(sourceName: string): string {
  const stem = sourceName.replace(/\.[^.]+$/, "").trim() || "apresentacao";
  return `${stem}.pptx`;
}

function sniffCharset(html: string): string | null {
  const charsetMatch = html.match(/<meta[^>]+charset=["']?\s*([a-z0-9\-_]+)/i);
  if (charsetMatch?.[1]) {
    return charsetMatch[1].trim();
  }

  const contentMatch = html.match(/content=["'][^"']*charset=([a-z0-9\-_]+)/i);
  return contentMatch?.[1]?.trim() ?? null;
}

function hasRelativeAssets(documentNode: Document): boolean {
  return Array.from(
    documentNode.querySelectorAll(
      "img[src], source[src], video[src], audio[src], track[src], script[src], link[href][rel='stylesheet'], link[href][rel='preload'], link[href][rel='icon']"
    )
  ).some((element) => {
    if (!(element instanceof HTMLElement)) {
      return false;
    }

    const value = element.getAttribute("src") || element.getAttribute("href");
    if (!value) {
      return false;
    }

    return isRelativeAsset(value);
  });
}

function hasGoogleFontsWithoutCrossOrigin(documentNode: Document): boolean {
  return Array.from(documentNode.querySelectorAll('link[href*="fonts.googleapis.com"]')).some((element) => {
    if (!(element instanceof HTMLLinkElement)) {
      return false;
    }

    return !element.crossOrigin;
  });
}

function isRelativeAsset(value: string): boolean {
  if (
    value.startsWith("#") ||
    value.startsWith("data:") ||
    value.startsWith("blob:") ||
    value.startsWith("http://") ||
    value.startsWith("https://") ||
    value.startsWith("//") ||
    value.startsWith("mailto:") ||
    value.startsWith("tel:")
  ) {
    return false;
  }

  return !/^[a-zA-Z]:\\/.test(value) && !value.startsWith("/");
}
