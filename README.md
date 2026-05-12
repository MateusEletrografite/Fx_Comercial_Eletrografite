# Fx Comercial Eletrografite

Dashboard comercial estatico para GitHub Pages.

## Publicacao

O arquivo principal do site e `index.html`. Ele foi gerado a partir de `01 - Apresentacao/01_DASHBOARD_ATUAL.html` e ajustado para rodar a partir da raiz do GitHub Pages.

## Conteudo publicado

- `index.html`: dashboard completo com dados embutidos.
- `05 - Apoio/digisac_supplement_matches.js`: complemento de dados usado pelo dashboard.
- `04 - Fonte/*.xlsx` e `04 - Fonte/*.pdf`: documentacao, bases consolidadas e guias.
- `08 - Codigos`: scripts usados para gerar e auditar o pacote.
- `09 - Regras`: regras e mapa mental.

## Observacao operacional

As bases CSV brutas e historicos grandes ficaram fora do Git para evitar falhas no GitHub, especialmente o limite de 100 MB por arquivo. Para atualizar a base do dashboard, gere novamente o HTML e substitua o `index.html`.
