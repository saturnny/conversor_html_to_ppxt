# HTML para PowerPoint Editavel

Aplicacao web em Next.js para transformar arquivos HTML em apresentacoes PowerPoint (`.pptx`) baixaveis, priorizando elementos editaveis com `dom-to-pptx` e sem usar Python.

## Stack

- Next.js com App Router
- React + TypeScript
- Tailwind CSS
- `dom-to-pptx` para exportacao editavel no navegador
- API Routes do Next.js apenas para servir o bundle do exportador e manter compatibilidade com Vercel

## Como funciona

1. O usuario faz upload de um arquivo `.html` ou `.htm`.
2. O navegador le o arquivo e monta um sandbox isolado com o HTML renderizado.
3. O app detecta os containers de slide mais provaveis, incluindo decks com `.slide`.
4. O `dom-to-pptx` percorre o DOM renderizado e gera um `.pptx` com textos, caixas, imagens e SVGs editaveis sempre que houver suporte.
5. O arquivo final eh baixado no navegador.

## O que esta coberto

- Upload de HTML
- Deteccao automatica de multiplos slides
- Exportacao editavel baseada no DOM renderizado
- Suporte a SVG como vetor
- Incorporacao automatica de fontes quando o ambiente permitir
- Feedback visual durante leitura, preparo e exportacao
- Interface desktop-first pronta para deploy na Vercel

## Limites importantes

- HTML e CSS muito customizados ainda podem exigir ajustes, principalmente quando dependem de JavaScript, `transform`, `backdrop-filter` ou assets locais.
- Scripts do HTML sao removidos por seguranca antes da exportacao.
- Imagens, fontes e folhas de estilo externas precisam responder com CORS adequado.
- Arquivos referenciados por caminho local ou relativo precisam estar embutidos no HTML ou disponiveis em URL publica.

## Rodando localmente

```bash
npm install
npm run dev
```

Abra [http://localhost:3000](http://localhost:3000).

## Scripts

- `npm run dev`: ambiente local
- `npm run build`: build de producao
- `npm run start`: sobe o build pronto
- `npm run typecheck`: validacao TypeScript

## Deploy na Vercel

1. Suba este projeto para um repositorio Git.
2. Importe o repositorio na Vercel.
3. A Vercel detecta automaticamente o projeto Next.js.
4. Faca o deploy sem precisar de servidor adicional.

## Componente reutilizavel

O projeto tambem inclui um componente pronto em JavaScript/React para exportacao manual:

- `components/exportar-apresentacao.jsx`

Ele aceita seletores CSS, elementos HTML ou refs React e chama `exportToPptx` com `async/await`, `autoEmbedFonts`, `svgAsVector` e tratamento de erro.
