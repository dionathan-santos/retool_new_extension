# Retool Suite Importer (Chrome Extension)

Esta extensão empacota o script já funcional para execução com 1 clique no Chrome.

## Como usar

1. Abra `chrome://extensions`.
2. Ative **Developer mode**.
3. Clique em **Load unpacked** e selecione esta pasta.
4. Abra a página do Retool (a tela onde você já rodava o script no console).
5. Clique no ícone da extensão **Retool Suite Importer**.
6. O mesmo fluxo do script original será iniciado (SheetJS, seletor de arquivo, preenchimento).

## Observações

- A lógica do script foi mantida; apenas foi empacotada como extensão.
- Para contornar CSP da página, a extensão tenta carregar SheetJS por `import()` primeiro (sem injetar `<script>` no DOM); se o navegador não suportar, usa fallback por `<script src>`.
- A extensão precisa acesso de rede ao domínio `https://cdn.sheetjs.com` para carregar o SheetJS.
