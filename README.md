# Regulamento — Programa Ideia Premiada (Arrais Advogados)

Gerador automático do **Regulamento Interno do Programa Ideia Premiada** do escritório
Arrais Advogados. O script `gerar_regulamento.js` monta o documento em formato Word
(`.docx`) com a identidade visual do escritório (fonte Arial, títulos em azul-marinho,
numeração de páginas e tabelas formatadas).

## Pré-requisitos

- [Node.js](https://nodejs.org/) instalado (versão 16 ou superior).

## Como usar

1. Instale as dependências:

   ```bash
   npm install
   ```

2. Gere o documento:

   ```bash
   npm start
   ```

   ou, diretamente:

   ```bash
   node gerar_regulamento.js
   ```

3. O arquivo `.docx` será criado na mesma pasta do script.

## Estrutura

- `gerar_regulamento.js` — script que monta e exporta o regulamento em `.docx`.
- `package.json` — metadados e dependências do projeto.

## Dependência

- [`docx`](https://www.npmjs.com/package/docx) — biblioteca para geração de arquivos Word.
