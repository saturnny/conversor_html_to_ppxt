import { existsSync } from "node:fs";
import chromium from "@sparticuz/chromium";
import { chromium as playwrightChromium } from "playwright-core";
import type { Browser, Page } from "playwright-core";
import type { RenderedPresentation, RenderedSlide } from "../presentation-types";

const VIEWPORT = {
  width: 1920,
  height: 1080
};

const EXPLICIT_SLIDE_SELECTOR =
  ".presentation-container > .slide, body > .slide, [data-slide], [data-ppt-slide], .ppt-slide, .page-slide";

const LOCAL_BROWSER_CANDIDATES = [
  process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH,
  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
  "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
].filter((value): value is string => Boolean(value));

export async function renderHtmlToSlides(
  html: string,
  sourceName: string
): Promise<RenderedPresentation> {
  const browser = await launchBrowser();

  try {
    const page = await browser.newPage({
      viewport: VIEWPORT,
      deviceScaleFactor: 1
    });

    await page.route("**/*", async (route) => {
      const url = route.request().url();

      if (isAllowedRequest(url)) {
        await route.continue();
        return;
      }

      await route.abort();
    });

    await page.setContent(sanitizeHtml(extractRenderableHtml(html)), {
      waitUntil: "load",
      timeout: 30000
    });

    await page.emulateMedia({ media: "screen" });
    await page.waitForLoadState("networkidle", { timeout: 10000 }).catch(() => undefined);
    await page.addStyleTag({
      content: `
        * {
          animation-duration: 0s !important;
          animation-delay: 0s !important;
          transition-duration: 0s !important;
          transition-delay: 0s !important;
          caret-color: transparent !important;
        }
      `
    });

    const warnings: string[] = [];
    const slideCount = await page.locator(EXPLICIT_SLIDE_SELECTOR).count();

    const slides =
      slideCount > 1
        ? await renderExplicitSlides(page, slideCount)
        : await renderFallbackSlide(page, warnings);

    return {
      fileName: toOutputFileName(sourceName),
      sourceName,
      slides,
      warnings
    };
  } finally {
    await browser.close();
  }
}

async function renderExplicitSlides(page: Page, totalSlides: number): Promise<RenderedSlide[]> {
  const slides: RenderedSlide[] = [];

  for (let index = 0; index < totalSlides; index += 1) {
    await prepareExplicitSlide(page, index, totalSlides);
    const locator = page.locator(EXPLICIT_SLIDE_SELECTOR).nth(index);

    await locator.scrollIntoViewIfNeeded();
    await page.waitForTimeout(50);

    const box = await locator.boundingBox();
    if (!box) {
      continue;
    }

    const title = await locator.evaluate((element) => {
      const titleElement = element.querySelector("[data-slide-title], h1, h2, h3");
      return titleElement?.textContent?.trim() || `Slide`;
    });

    const imageBuffer = await locator.screenshot({
      type: "png",
      animations: "disabled"
    });

    slides.push({
      title: title || `Slide ${index + 1}`,
      width: Math.round(box.width),
      height: Math.round(box.height),
      imageDataUrl: toPngDataUrl(imageBuffer)
    });
  }

  return slides;
}

async function prepareExplicitSlide(page: Page, index: number, totalSlides: number): Promise<void> {
  await page.evaluate(
    ({ selector, targetIndex, total }) => {
      const slides = Array.from(document.querySelectorAll<HTMLElement>(selector));

      slides.forEach((slide, slideIndex) => {
        const isActive = slideIndex === targetIndex;
        slide.classList.toggle("active", isActive);
        slide.classList.add("print-visible");

        slide.style.setProperty("visibility", isActive ? "visible" : "hidden", "important");
        slide.style.setProperty("opacity", isActive ? "1" : "0", "important");
        slide.style.setProperty("display", "flex", "important");
        slide.style.setProperty("z-index", isActive ? "999" : "0", "important");
        slide.style.setProperty("pointer-events", "none", "important");

        slide.querySelectorAll<HTMLElement>(".page-indicator").forEach((indicator) => {
          indicator.textContent = `${targetIndex + 1} / ${total}`;
        });

        slide.querySelectorAll<HTMLElement>(".footer-progress-bar").forEach((progressBar) => {
          progressBar.style.setProperty("width", `${((targetIndex + 1) / total) * 100}%`, "important");
        });
      });

      document.querySelectorAll<HTMLElement>(".nav-btn").forEach((button) => {
        button.style.setProperty("display", "none", "important");
      });
    },
    {
      selector: EXPLICIT_SLIDE_SELECTOR,
      targetIndex: index,
      total: totalSlides
    }
  );
}

async function renderFallbackSlide(page: Page, warnings: string[]): Promise<RenderedSlide[]> {
  warnings.push(
    "O HTML nao expunha slides separados no DOM. Foi capturada uma imagem unica da pagina como fallback."
  );

  const bodyLocator = page.locator("body");
  const bodyBox = await bodyLocator.boundingBox();
  const screenshot = await page.screenshot({
    type: "png",
    fullPage: true,
    animations: "disabled"
  });

  return [
    {
      title: "Slide 1",
      width: Math.round(bodyBox?.width || VIEWPORT.width),
      height: Math.round(bodyBox?.height || VIEWPORT.height),
      imageDataUrl: toPngDataUrl(screenshot)
    }
  ];
}

async function launchBrowser(): Promise<Browser> {
  const localExecutable = LOCAL_BROWSER_CANDIDATES.find((candidate) => existsOnCurrentMachine(candidate));

  if (localExecutable) {
    return playwrightChromium.launch({
      executablePath: localExecutable,
      headless: true
    });
  }

  return playwrightChromium.launch({
    executablePath: await chromium.executablePath(),
    args: chromium.args,
    headless: true
  });
}

function existsOnCurrentMachine(filePath: string): boolean {
  try {
    return existsSync(filePath);
  } catch {
    return false;
  }
}

function sanitizeHtml(html: string): string {
  return html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
    .replace(/\son[a-z]+\s*=\s*(['"]).*?\1/gi, "");
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

function isAllowedRequest(url: string): boolean {
  try {
    const protocol = new URL(url).protocol;
    return protocol === "http:" || protocol === "https:" || protocol === "data:" || protocol === "about:";
  } catch {
    return false;
  }
}

function toPngDataUrl(buffer: Buffer): string {
  return `data:image/png;base64,${buffer.toString("base64")}`;
}

function toOutputFileName(sourceName: string): string {
  const stem = sourceName.replace(/\.[^.]+$/, "").trim() || "apresentacao";
  return `${stem}.pptx`;
}
