"use client";

import { useEffect, useRef, useState } from "react";
import {
  EXPORT_FRAME_API_KEY,
  type ExportFrameApi,
  type ExportFrameWindow,
  decodeHtmlFile,
  prepareEditableExportDocument
} from "../lib/client/editable-export";

type ConverterStatus = "idle" | "reading" | "preparing" | "ready" | "exporting" | "success" | "error";

type PreparedSandbox = ReturnType<typeof prepareEditableExportDocument> & {
  sessionId: number;
};

const STATUS_LABELS: Record<ConverterStatus, string> = {
  idle: "Envie um HTML para iniciar",
  reading: "Lendo o HTML enviado",
  preparing: "Preparando o sandbox editável",
  ready: "Pronto para exportar",
  exporting: "Exportando para PowerPoint",
  success: "Arquivo gerado com sucesso",
  error: "Falha durante a exportação"
};

const DEFAULT_LIST_CONFIG = {
  spacing: {
    before: 6,
    after: 3
  }
};

export function ConverterWorkspace() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const iframeRef = useRef<HTMLIFrameElement | null>(null);
  const sessionRef = useRef(0);

  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [preparedSandbox, setPreparedSandbox] = useState<PreparedSandbox | null>(null);
  const [status, setStatus] = useState<ConverterStatus>("idle");
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [slideCount, setSlideCount] = useState<number>(0);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState<string>("apresentacao.pptx");
  const [lastSourceName, setLastSourceName] = useState<string>("Nenhum arquivo selecionado");

  useEffect(() => {
    return () => {
      if (downloadUrl) {
        URL.revokeObjectURL(downloadUrl);
      }
    };
  }, [downloadUrl]);

  const isBusy = status === "reading" || status === "preparing" || status === "exporting";
  const isSuccess = status === "success" && Boolean(downloadUrl);
  const canExport = Boolean(preparedSandbox && selectedFile) && !isBusy && status !== "error";
  const statusLabel = STATUS_LABELS[status];

  async function processSelectedFile(file: File | null) {
    const sessionId = sessionRef.current + 1;
    sessionRef.current = sessionId;

    clearDownloadUrl();
    setSelectedFile(null);
    setPreparedSandbox(null);
    setWarnings([]);
    setSlideCount(0);
    setError(null);

    if (!file) {
      setLastSourceName("Nenhum arquivo selecionado");
      setDownloadName("apresentacao.pptx");
      setStatus("idle");
      return;
    }

    setLastSourceName(file.name);

    if (!isHtmlFile(file)) {
      setDownloadName("apresentacao.pptx");
      setError("Selecione um arquivo .html ou .htm para continuar.");
      setStatus("error");
      return;
    }

    setSelectedFile(file);
    setStatus("reading");

    try {
      const html = await decodeHtmlFile(file);

      if (sessionRef.current !== sessionId) {
        return;
      }

      const prepared = prepareEditableExportDocument(html, file.name);
      setPreparedSandbox({
        ...prepared,
        sessionId
      });
      setDownloadName(prepared.downloadName);
      setWarnings(prepared.warnings);
      setStatus("preparing");
    } catch (processingError) {
      if (sessionRef.current !== sessionId) {
        return;
      }

      const message =
        processingError instanceof Error
          ? processingError.message
          : "Nao foi possivel carregar o HTML enviado.";

      setSelectedFile(null);
      setPreparedSandbox(null);
      setError(message);
      setStatus("error");
    }
  }

  async function handleSandboxLoad() {
    if (!preparedSandbox) {
      return;
    }

    const activeSession = preparedSandbox.sessionId;

    try {
      const api = await waitForFrameApi(iframeRef.current);
      const result = await api.inspect();

      if (sessionRef.current !== activeSession) {
        return;
      }

      setSlideCount(result.slideCount);
      setWarnings(uniqueStrings([...preparedSandbox.warnings, ...result.warnings]));
      setStatus("ready");
    } catch (sandboxError) {
      if (sessionRef.current !== activeSession) {
        return;
      }

      const message =
        sandboxError instanceof Error
          ? sandboxError.message
          : "O sandbox de exportacao nao conseguiu preparar os slides.";

      setError(message);
      setStatus("error");
    }
  }

  async function handleExport() {
    if (!preparedSandbox || !selectedFile) {
      setError("Selecione um arquivo HTML antes de exportar.");
      setStatus("error");
      return;
    }

    const activeSession = preparedSandbox.sessionId;
    setError(null);
    setStatus("exporting");

    try {
      const api = await waitForFrameApi(iframeRef.current);
      const result = await api.exportSlides({
        fileName: preparedSandbox.downloadName,
        autoEmbedFonts: true,
        fonts: [],
        svgAsVector: true,
        listConfig: DEFAULT_LIST_CONFIG
      });

      if (sessionRef.current !== activeSession) {
        return;
      }

      const nextUrl = URL.createObjectURL(result.blob);
      setDownloadUrl(nextUrl);
      setDownloadName(result.fileName || preparedSandbox.downloadName);
      setSlideCount(result.slideCount);
      setWarnings(uniqueStrings([...preparedSandbox.warnings, ...result.warnings]));
      setStatus("success");

      triggerDownload(nextUrl, result.fileName || preparedSandbox.downloadName);
    } catch (exportError) {
      if (sessionRef.current !== activeSession) {
        return;
      }

      const message = formatExportError(exportError);
      setError(message);
      setStatus("error");
    }
  }

  function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    void processSelectedFile(event.target.files?.[0] ?? null);
  }

  function handleDragOver(event: React.DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    if (!isBusy) {
      setIsDragging(true);
    }
  }

  function handleDragLeave(event: React.DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
      setIsDragging(false);
    }
  }

  function handleDrop(event: React.DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    setIsDragging(false);

    if (isBusy) {
      return;
    }

    void processSelectedFile(event.dataTransfer.files?.[0] ?? null);
  }

  function handleReset() {
    sessionRef.current += 1;
    clearDownloadUrl();

    if (inputRef.current) {
      inputRef.current.value = "";
    }

    setSelectedFile(null);
    setPreparedSandbox(null);
    setLastSourceName("Nenhum arquivo selecionado");
    setWarnings([]);
    setSlideCount(0);
    setError(null);
    setDownloadName("apresentacao.pptx");
    setStatus("idle");
  }

  function clearDownloadUrl() {
    setDownloadUrl((currentUrl) => {
      if (currentUrl) {
        URL.revokeObjectURL(currentUrl);
      }

      return null;
    });
  }

  return (
    <>
      {preparedSandbox ? (
        <iframe
          key={preparedSandbox.sessionId}
          ref={iframeRef}
          title="Sandbox de exportacao editavel"
          sandbox="allow-same-origin allow-scripts"
          srcDoc={preparedSandbox.srcDoc}
          onLoad={handleSandboxLoad}
          className="pointer-events-none fixed -left-[200vw] top-0 h-[1080px] w-[1920px] opacity-0"
        />
      ) : null}

      <section className="mx-auto w-full max-w-4xl overflow-hidden rounded-[28px] border border-slate-200 bg-white shadow-[0_24px_90px_rgba(15,23,42,0.08)]">
        <div className="flex items-center justify-between border-b border-slate-200 px-8 py-5">
          <div>
            <p className="text-sm font-semibold text-slate-950">Upload e exportacao editavel</p>
            <p className="mt-1 text-sm text-slate-500">
              O HTML eh renderizado em um sandbox isolado e exportado com elementos editaveis no PowerPoint.
            </p>
          </div>

          <div className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-slate-50 px-4 py-2 text-sm font-medium text-slate-600">
            <span
              className={`h-2 w-2 rounded-full ${
                isSuccess ? "bg-emerald-500" : status === "error" ? "bg-rose-500" : "bg-blue-500"
              }`}
            />
            {statusLabel}
          </div>
        </div>

        {isSuccess ? (
          <div className="flex min-h-[470px] flex-col items-center justify-center px-10 py-14 text-center">
            <div className="flex h-20 w-20 items-center justify-center rounded-full bg-emerald-50 text-emerald-600">
              <SuccessIcon />
            </div>

            <h2 className="mt-8 text-3xl font-semibold tracking-[-0.03em] text-slate-950">
              Arquivo convertido com sucesso
            </h2>

            <p className="mt-3 max-w-xl text-base leading-7 text-slate-600">
              O download do PPTX ja foi iniciado. {slideCount > 0 ? `${slideCount} slides editaveis foram preparados.` : ""}
            </p>

            {warnings.length > 0 ? (
              <WarningPanel
                title="Pontos de atencao detectados"
                warnings={warnings}
                className="mt-6 w-full max-w-2xl text-left"
              />
            ) : null}

            <div className="mt-8 flex items-center gap-4">
              <a
                className="inline-flex h-14 items-center justify-center rounded-xl bg-blue-600 px-7 text-base font-semibold text-white shadow-[0_14px_32px_rgba(37,99,235,0.24)] transition hover:bg-blue-500"
                href={downloadUrl ?? "#"}
                download={downloadName}
              >
                Baixar arquivo PPTX
              </a>
              <button
                type="button"
                onClick={handleReset}
                className="text-sm font-medium text-slate-600 transition hover:text-slate-950"
              >
                Converter outro arquivo
              </button>
            </div>
          </div>
        ) : (
          <div className="px-8 py-8">
            <label
              htmlFor="html-file"
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`group flex min-h-[240px] cursor-pointer flex-col items-center justify-center rounded-3xl border-2 border-dashed px-8 text-center transition ${
                isDragging
                  ? "border-blue-500 bg-blue-50/80"
                  : "border-slate-300 bg-slate-50/80 hover:border-blue-400 hover:bg-blue-50/60"
              }`}
            >
              <input
                ref={inputRef}
                id="html-file"
                className="sr-only"
                type="file"
                accept=".html,.htm,text/html"
                onChange={handleFileChange}
              />

              <div className="flex h-20 w-20 items-center justify-center rounded-3xl bg-white text-blue-600 shadow-sm ring-1 ring-slate-200 transition group-hover:scale-[1.02]">
                <UploadIcon />
              </div>

              <div className="mt-6">
                <p className="text-xl font-semibold text-slate-950">
                  Arraste seu HTML ou clique para selecionar
                </p>
                <p className="mt-2 text-sm leading-6 text-slate-500">
                  O arquivo eh carregado no navegador e preparado para exportacao editavel com{" "}
                  <span className="font-medium text-slate-700">dom-to-pptx</span>.
                </p>
              </div>
            </label>

            <div className="mt-5 grid gap-4 rounded-2xl border border-slate-200 bg-slate-50 px-5 py-4 md:grid-cols-[minmax(0,1fr)_auto_auto] md:items-center">
              <div className="min-w-0">
                <p className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">
                  Arquivo selecionado
                </p>
                <p className="mt-1 truncate text-sm font-medium text-slate-900">{lastSourceName}</p>
              </div>

              <div className="rounded-full border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600">
                HTML to PPTX editavel
              </div>

              <div className="rounded-full border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600">
                {slideCount > 0 ? `${slideCount} slides detectados` : "Aguardando preparo"}
              </div>
            </div>

            <div className="mt-8 flex flex-col items-center">
              <button
                className="inline-flex h-14 min-w-[320px] items-center justify-center rounded-xl bg-blue-600 px-7 text-base font-semibold text-white shadow-[0_16px_36px_rgba(37,99,235,0.28)] transition hover:bg-blue-500 disabled:cursor-not-allowed disabled:bg-blue-300 disabled:shadow-none"
                type="button"
                disabled={!canExport}
                onClick={handleExport}
              >
                {isBusy ? (
                  <span className="inline-flex items-center gap-3">
                    <SpinnerIcon />
                    {status === "reading"
                      ? "Lendo HTML..."
                      : status === "preparing"
                        ? "Preparando slides..."
                        : "Exportando..."}
                  </span>
                ) : (
                  "Exportar para PowerPoint"
                )}
              </button>

              <p className="mt-4 text-sm text-slate-500">
                Quando o sandbox termina de preparar o HTML, o botao libera a exportacao editavel.
              </p>
            </div>

            {error ? (
              <div className="mt-6 rounded-2xl border border-rose-200 bg-rose-50 px-5 py-4 text-sm text-rose-700" role="alert">
                {error}
              </div>
            ) : null}

            {warnings.length > 0 ? (
              <WarningPanel title="Pontos de atencao" warnings={warnings} className="mt-4" />
            ) : null}
          </div>
        )}
      </section>
    </>
  );
}

type WarningPanelProps = {
  className?: string;
  title: string;
  warnings: string[];
};

function WarningPanel({ className, title, warnings }: WarningPanelProps) {
  return (
    <div className={`rounded-2xl border border-amber-200 bg-amber-50 px-5 py-4 text-sm text-amber-900 ${className ?? ""}`}>
      <p className="font-semibold">{title}</p>
      <ul className="mt-2 list-disc space-y-1 pl-5">
        {warnings.map((warning) => (
          <li key={warning}>{warning}</li>
        ))}
      </ul>
    </div>
  );
}

async function waitForFrameApi(frame: HTMLIFrameElement | null): Promise<ExportFrameApi> {
  const timeoutAt = Date.now() + 15000;

  while (Date.now() < timeoutAt) {
    const frameWindow = frame?.contentWindow as ExportFrameWindow | null;
    const api = frameWindow?.[EXPORT_FRAME_API_KEY];

    if (api) {
      return api;
    }

    await wait(120);
  }

  throw new Error("O sandbox de exportacao demorou demais para inicializar.");
}

function wait(delayMs: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, delayMs);
  });
}

function triggerDownload(url: string, fileName: string) {
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
}

function formatExportError(error: unknown): string {
  if (error instanceof Error) {
    if (/cors|tainted|font|image/i.test(error.message)) {
      return `${error.message} Verifique CORS de imagens e fontes externas usadas no HTML.`;
    }

    return error.message;
  }

  return "Nao foi possivel exportar o PowerPoint editavel.";
}

function uniqueStrings(values: string[]): string[] {
  return Array.from(new Set(values.filter(Boolean)));
}

function isHtmlFile(file: File): boolean {
  return file.type === "text/html" || /\.(html?|HTML?)$/.test(file.name);
}

function UploadIcon() {
  return (
    <svg aria-hidden="true" className="h-9 w-9" viewBox="0 0 24 24" fill="none">
      <path
        d="M12 16V6M12 6l-4 4M12 6l4 4M5 18.5h14"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function SpinnerIcon() {
  return (
    <svg aria-hidden="true" className="h-5 w-5 animate-spin" viewBox="0 0 24 24" fill="none">
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeOpacity="0.24" strokeWidth="3" />
      <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="3" strokeLinecap="round" />
    </svg>
  );
}

function SuccessIcon() {
  return (
    <svg aria-hidden="true" className="h-9 w-9" viewBox="0 0 24 24" fill="none">
      <path
        d="M7 12.5l3.2 3.2L17.5 8.5"
        stroke="currentColor"
        strokeWidth="2.2"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}
