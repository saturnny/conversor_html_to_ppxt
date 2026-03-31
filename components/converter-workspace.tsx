"use client";

import { useEffect, useMemo, useRef, useState } from "react";

type ConverterStatus = "idle" | "uploading" | "rendering" | "generating" | "success" | "error";

const STATUS_LABELS: Record<ConverterStatus, string> = {
  idle: "Pronto para receber um HTML",
  uploading: "Enviando o HTML para processamento",
  rendering: "Renderizando os slides no backend",
  generating: "Gerando o PowerPoint editável",
  success: "Arquivo gerado com sucesso",
  error: "Falha durante a conversão"
};

export function ConverterWorkspace() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
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

  const isBusy = status === "uploading" || status === "rendering" || status === "generating";
  const isSuccess = status === "success" && Boolean(downloadUrl);
  const statusLabel = useMemo(() => STATUS_LABELS[status], [status]);

  async function handleSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (!selectedFile) {
      setError("Selecione um arquivo .html ou .htm para iniciar a conversão.");
      setStatus("error");
      return;
    }

    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      setDownloadUrl(null);
    }

    setError(null);
    setWarnings([]);
    setSlideCount(0);
    setLastSourceName(selectedFile.name);

    try {
      setStatus("uploading");

      const formData = new FormData();
      formData.append("htmlFile", selectedFile);

      setStatus("rendering");

      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData
      });

      if (!response.ok) {
        const errorMessage = await extractErrorMessage(response);
        throw new Error(errorMessage);
      }

      const reportedSlideCount = Number(response.headers.get("x-slide-count") || "0");
      const warningsHeader = response.headers.get("x-conversion-warnings");

      if (reportedSlideCount > 0) {
        setSlideCount(reportedSlideCount);
      }

      if (warningsHeader) {
        try {
          setWarnings(JSON.parse(warningsHeader) as string[]);
        } catch {
          setWarnings([]);
        }
      }

      setStatus("generating");

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);

      setDownloadUrl(url);
      setDownloadName(extractDownloadName(response.headers.get("content-disposition"), selectedFile.name));
      setStatus("success");
    } catch (conversionError) {
      const message =
        conversionError instanceof Error
          ? conversionError.message
          : "Não foi possível converter o HTML em PowerPoint.";

      setError(message);
      setStatus("error");
    }
  }

  function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    setFileSelection(event.target.files?.[0] ?? null);
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

    const file = event.dataTransfer.files?.[0] ?? null;
    setFileSelection(file);
  }

  function handleReset() {
    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      setDownloadUrl(null);
    }

    if (inputRef.current) {
      inputRef.current.value = "";
    }

    setSelectedFile(null);
    setLastSourceName("Nenhum arquivo selecionado");
    setWarnings([]);
    setSlideCount(0);
    setError(null);
    setStatus("idle");
  }

  function setFileSelection(file: File | null) {
    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      setDownloadUrl(null);
    }

    if (!file) {
      setSelectedFile(null);
      setLastSourceName("Nenhum arquivo selecionado");
      setWarnings([]);
      setSlideCount(0);
      setError(null);
      setStatus("idle");
      return;
    }

    if (!isHtmlFile(file)) {
      setSelectedFile(null);
      setLastSourceName(file.name);
      setWarnings([]);
      setSlideCount(0);
      setError("Selecione um arquivo .html ou .htm para continuar.");
      setStatus("error");
      return;
    }

    setSelectedFile(file);
    setLastSourceName(file.name);
    setWarnings([]);
    setSlideCount(0);
    setError(null);
    setStatus("idle");
  }

  return (
    <form
      onSubmit={handleSubmit}
      className="mx-auto w-full max-w-4xl overflow-hidden rounded-[28px] border border-slate-200 bg-white shadow-[0_24px_90px_rgba(15,23,42,0.08)]"
    >
      <div className="flex items-center justify-between border-b border-slate-200 px-8 py-5">
        <div>
          <p className="text-sm font-semibold text-slate-950">Upload e conversão</p>
          <p className="mt-1 text-sm text-slate-500">
            Um único fluxo para enviar, converter e baixar seu PowerPoint.
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
            Seu arquivo PPTX está pronto para download. {slideCount > 0 ? `${slideCount} slides foram preparados.` : ""}
          </p>

          {warnings.length > 0 ? (
            <div className="mt-6 w-full max-w-2xl rounded-2xl border border-amber-200 bg-amber-50 px-5 py-4 text-left text-sm text-amber-900">
              <p className="font-semibold">Pontos de atenção detectados</p>
              <ul className="mt-2 list-disc space-y-1 pl-5">
                {warnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            </div>
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
                Compatível com arquivos <span className="font-medium text-slate-700">.html</span> e{" "}
                <span className="font-medium text-slate-700">.htm</span>.
              </p>
            </div>
          </label>

          <div className="mt-5 flex items-center justify-between gap-6 rounded-2xl border border-slate-200 bg-slate-50 px-5 py-4">
            <div className="min-w-0">
              <p className="text-xs font-semibold uppercase tracking-[0.18em] text-slate-500">
                Arquivo selecionado
              </p>
              <p className="mt-1 truncate text-sm font-medium text-slate-900">{lastSourceName}</p>
            </div>

            <div className="rounded-full border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600">
              HTML → PPTX
            </div>
          </div>

          <div className="mt-8 flex flex-col items-center">
            <button
              className="inline-flex h-14 min-w-[320px] items-center justify-center rounded-xl bg-blue-600 px-7 text-base font-semibold text-white shadow-[0_16px_36px_rgba(37,99,235,0.28)] transition hover:bg-blue-500 disabled:cursor-not-allowed disabled:bg-blue-300 disabled:shadow-none"
              type="submit"
              disabled={!selectedFile || isBusy}
            >
              {isBusy ? (
                <span className="inline-flex items-center gap-3">
                  <SpinnerIcon />
                  Convertendo...
                </span>
              ) : (
                "Converter para PowerPoint"
              )}
            </button>

            <p className="mt-4 text-sm text-slate-500">
              Tudo processado no ambiente do Next.js, com download imediato ao final.
            </p>
          </div>

          {error ? (
            <div className="mt-6 rounded-2xl border border-rose-200 bg-rose-50 px-5 py-4 text-sm text-rose-700" role="alert">
              {error}
            </div>
          ) : null}

          {warnings.length > 0 ? (
            <div className="mt-4 rounded-2xl border border-amber-200 bg-amber-50 px-5 py-4 text-sm text-amber-900">
              <p className="font-semibold">Pontos de atenção</p>
              <ul className="mt-2 list-disc space-y-1 pl-5">
                {warnings.map((warning) => (
                  <li key={warning}>{warning}</li>
                ))}
              </ul>
            </div>
          ) : null}
        </div>
      )}
    </form>
  );
}

async function extractErrorMessage(response: Response): Promise<string> {
  const contentType = response.headers.get("content-type") ?? "";

  if (contentType.includes("application/json")) {
    const payload = (await response.json()) as { error?: string };
    return payload.error || "A API não conseguiu gerar o arquivo .pptx.";
  }

  const text = await response.text();
  return text || "A API não conseguiu gerar o arquivo .pptx.";
}

function extractDownloadName(contentDisposition: string | null, fallbackSourceName: string): string {
  if (contentDisposition) {
    const match = contentDisposition.match(/filename="?([^"]+)"?/i);
    if (match?.[1]) {
      return match[1];
    }
  }

  const sourceStem = fallbackSourceName.replace(/\.[^.]+$/, "") || "apresentacao";
  return `${sourceStem}.pptx`;
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
