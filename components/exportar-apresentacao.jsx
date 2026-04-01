"use client";

import { useMemo, useState } from "react";
import { exportToPptx } from "dom-to-pptx";

function resolverAlvos(slides) {
  const lista = Array.isArray(slides) ? slides : [slides];

  return lista
    .map((item) => {
      if (!item) {
        return null;
      }

      if (typeof item === "string") {
        return item;
      }

      if (item.current instanceof HTMLElement) {
        return item.current;
      }

      if (item instanceof HTMLElement) {
        return item;
      }

      return null;
    })
    .filter(Boolean);
}

function baixarBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export function ExportarApresentacao({
  slides,
  fileName = "apresentacao.pptx",
  fonts = [
    {
      name: "Roboto",
      url: "https://fonts.gstatic.com/s/roboto/v30/KFOmCnqEu92Fr1Mu4mxK.woff2"
    }
  ],
  svgAsVector = true,
  listConfig = {
    spacing: {
      before: 6,
      after: 3
    }
  },
  disabled = false,
  onSuccess,
  onError
}) {
  const [exportando, setExportando] = useState(false);
  const [erro, setErro] = useState("");
  const alvos = useMemo(() => resolverAlvos(slides), [slides]);

  async function handleExportar() {
    if (!alvos.length) {
      const mensagem = "Nenhum slide valido foi encontrado para exportacao.";
      setErro(mensagem);
      onError?.(mensagem);
      return;
    }

    setExportando(true);
    setErro("");

    try {
      // A biblioteca retorna um Blob, entao usamos async/await e fazemos o download manual.
      const blob = await exportToPptx(alvos, {
        fileName,
        autoEmbedFonts: true,
        fonts,
        svgAsVector,
        listConfig,
        skipDownload: true
      });

      baixarBlob(blob, fileName);
      onSuccess?.(blob);
    } catch (error) {
      const mensagem =
        error instanceof Error
          ? `${error.message} Verifique CORS de imagens e fontes externas usadas no HTML.`
          : "Nao foi possivel exportar a apresentacao.";

      setErro(mensagem);
      onError?.(mensagem);
    } finally {
      setExportando(false);
    }
  }

  return (
    <div className="flex flex-col items-center gap-3">
      <button
        type="button"
        onClick={handleExportar}
        disabled={disabled || exportando}
        className="inline-flex h-12 items-center justify-center rounded-xl bg-blue-600 px-6 text-sm font-semibold text-white transition hover:bg-blue-500 disabled:cursor-not-allowed disabled:bg-blue-300"
      >
        {exportando ? "Exportando..." : "Exportar para PowerPoint"}
      </button>

      {erro ? <p className="text-sm text-rose-600">{erro}</p> : null}
    </div>
  );
}
