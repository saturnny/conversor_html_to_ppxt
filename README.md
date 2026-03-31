# HTML para PowerPoint Editavel

Aplicacao web em Next.js para transformar arquivos HTML em apresentacoes PowerPoint (`.pptx`) editaveis e baixaveis, sem usar Python e sem backend externo.

## Stack

- Next.js com App Router
- React + TypeScript
- API Route serverless no proprio Next.js
- `pptxgenjs` para gerar o arquivo `.pptx`

## Como funciona

1. O usuario faz upload de um arquivo `.html` ou `.htm`.
2. O navegador renderiza esse HTML em um iframe isolado.
3. A API serverless renderiza o HTML em um navegador headless.
4. Quando o documento ja possui slides no DOM, cada slide e capturado com alta fidelidade visual.
5. A rota serverless gera o `.pptx` com um slide por captura renderizada.
6. O navegador libera o download do arquivo final.

## O que esta coberto

- Upload de HTML
- Renderizacao real do HTML no backend
- Multiples slides a partir de decks que usam `.slide` ou estruturas equivalentes
- Captura visual fiel de layouts com CSS, logos, cards e componentes
- Feedback visual durante a conversao
- Download do `.pptx` gerado
- Estrutura pronta para deploy na Vercel

## Limites importantes

- O modo atual prioriza fidelidade visual. Isso significa que os slides podem sair como capturas renderizadas dentro do PowerPoint.
- Scripts do HTML sao removidos por seguranca antes da renderizacao.
- Imagens e assets locais referenciados por caminho relativo precisam estar embutidos no HTML como `data URL` ou acessiveis por URL publica.
- Fontes web podem ser substituidas pelo PowerPoint se nao estiverem instaladas no ambiente final.

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

1. Suba este projeto para um repositório Git.
2. Importe o repositório na Vercel.
3. A Vercel detecta automaticamente o projeto Next.js.
4. Faça o deploy sem precisar de servidor adicional.

## Observacao sobre fidelidade

Esta implementacao prioriza fidelidade visual para decks HTML reais. Para entregas em que a aparencia precisa ficar muito proxima do navegador, a captura renderizada tende a funcionar melhor do que tentar reconstruir tudo como formas e textos nativos do PowerPoint.
