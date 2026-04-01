import { ConverterWorkspace } from "../components/converter-workspace";
import { SiteHeader } from "../components/site-header";

export default function HomePage() {
  return (
    <>
      <SiteHeader />

      <main className="relative min-h-screen overflow-hidden bg-slate-50 pt-[72px]">
        <div className="pointer-events-none absolute inset-x-0 top-[72px] -z-10 h-[520px] bg-[radial-gradient(circle_at_top,_rgba(37,99,235,0.12),_transparent_52%),linear-gradient(180deg,_#ffffff_0%,_#f8fafc_72%)]" />
        <div className="pointer-events-none absolute inset-x-0 top-[120px] -z-10 mx-auto h-[420px] max-w-[1280px] rounded-full bg-[radial-gradient(circle,_rgba(148,163,184,0.12),_transparent_58%)] blur-3xl" />

        <div className="mx-auto flex min-h-[calc(100vh-72px)] w-full max-w-[1280px] flex-col justify-center px-6 py-8">
          <section className="mx-auto flex w-full max-w-5xl flex-col items-center">
            <div className="max-w-3xl text-center">
              <div className="inline-flex items-center gap-2 rounded-full border border-blue-100 bg-blue-50 px-4 py-2 text-sm font-medium text-blue-700">
                <span className="h-2 w-2 rounded-full bg-blue-600" />
                Exportacao editavel com foco total na acao principal
              </div>

              <h1 className="mt-7 text-balance text-5xl font-semibold tracking-[-0.04em] text-slate-950 xl:text-[3.9rem]">
                Converta HTML em PowerPoint editavel
              </h1>

              <p className="mx-auto mt-5 max-w-2xl text-lg leading-8 text-slate-600">
                Faca upload do seu HTML, prepare os slides em um sandbox seguro e baixe o arquivo
                <span className="font-medium text-slate-900"> .pptx </span>
                com textos, blocos e estruturas editaveis sempre que o layout permitir.
              </p>
            </div>

            <div className="mt-10 w-full">
              <ConverterWorkspace />
            </div>
          </section>
        </div>
      </main>
    </>
  );
}
