const NAV_ITEMS = [
  { label: "FAQ", href: "#faq" },
  { label: "Sobre", href: "#sobre" }
];

export function SiteHeader() {
  return (
    <header className="fixed inset-x-0 top-0 z-50 border-b border-slate-200/80 bg-white/90 backdrop-blur-xl">
      <div className="mx-auto flex h-[72px] w-full max-w-[1280px] items-center justify-between px-6">
        <a href="#" className="text-sm font-semibold tracking-[0.22em] text-slate-950 uppercase">
          HTML to PPTX
        </a>

        <nav className="flex items-center gap-3 text-sm font-medium text-slate-600">
          {NAV_ITEMS.map((item) => (
            <a
              key={item.label}
              href={item.href}
              className="rounded-full px-3 py-2 transition hover:bg-slate-100 hover:text-slate-950"
            >
              {item.label}
            </a>
          ))}

          <a
            href="mailto:contato@htmltopptx.app"
            className="ml-2 inline-flex h-10 items-center rounded-full bg-blue-600 px-4 text-white shadow-[0_12px_32px_rgba(37,99,235,0.24)] transition hover:bg-blue-500"
          >
            Contato
          </a>
        </nav>
      </div>
    </header>
  );
}
