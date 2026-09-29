# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-page executive dashboard ("Terra Biomassa") for the company's board, covering wood-chip biomass shipments, receivables, freight and driver payments. It is a static site with no build step, published on GitHub Pages, and its data is regenerated from an internal Excel workbook (currently **v13**). The published page carries the data **encrypted**; only the password decrypts it.

## Commands

There is no build/lint/test tooling: one HTML file plus one Python script.

```bash
# Regenerate from the newest .xlsx in dados/, encrypt, then git add + commit + push index.html
python gerar_dashboard.py

# Validate headers + reconcile against the workbook's PAINEL sheet, writing nothing
python gerar_dashboard.py --check

# Regenerate and encrypt, but skip git
python gerar_dashboard.py --sem-publicar
```

Requires `openpyxl` and `cryptography`. The password comes from the `TB_SENHA` env var, or is prompted (twice) on the terminal; it is never written to disk. `git` needs push access to `origin`.

Manual browser testing: `python -m http.server 8000` and open `dashboard.html` (plain data, no password) or `index.html` (encrypted, asks the password).

## Architecture

### The files — know which one to edit

- **`dashboard.html`**: the source to hand-edit (markup, CSS, JS). Gitignored. Its `const DATA = {...};` line holds the data in plain text for local development and `const PAYLOAD = null;`.
- **`index.html`**: generated from `dashboard.html` with `const DATA = null;` and `const PAYLOAD = {v,iter,salt,iv,ct};` (PBKDF2-SHA256 600k iterations → AES-256-GCM). The only data file committed. Never hand-edit.
- **`dashboard_slim.json`**, **`historico_kpis.json`**: local only (gitignored). The JSON is the plain data; the history keeps one KPI snapshot per workbook emission date and feeds the "evolução desde o relatório anterior" card (it was backfilled from the old plaintext commits on first run).

The generator replaces those two lines by prefix (`^const DATA = .*;$`, `^const PAYLOAD = .*;$`), so each must stay on its own line. Before writing `index.html` it aborts if any client or driver name would appear in plain text.

### `dashboard.html` structure

Self-contained: inline CSS/JS; external only Chart.js and Google Fonts from CDN. The logo is embedded once, as the favicon data URI; the `<img data-logo>` elements copy it at load time.

- **Access**: with `PAYLOAD`, the gate derives the key from the password and decrypts; a failed GCM tag means wrong password. The raw key is kept in `sessionStorage.tbKey` for the tab session. Without `PAYLOAD` (local), `DATA` opens directly.
- **Views**: seven, declared in the `VIEWS` array (id, titles, icon, charts). The sidebar and mobile bottom nav are generated from it by `montarNav()`; adding a view = add to `VIEWS` + a `<section class="view" id="view-<id>">` + a `render<View>()` called from `renderAll()`.
- **Filters → `recortes()` → `indicadores()`**: date semantics mirror the workbook's PAINEL: operational KPIs filter by `dataBase` (EMBARQUES col. 91), financial ones by `dataCompetencia` (col. 98, unload date), receipts by receipt date, "pago" by payment dates. No dates = accumulated, including undated rows. **`indicadores()` in JS and `indicadores()` in Python must stay identical**; the Python one is what `--check` reconciles against the PAINEL.
- **Reference date** for overdue/aging/"vence em N dias" is `meta.emissao` (the workbook snapshot), not the browser's today.
- **Drivers**: totals are computed from shipment rows (`motoristasApurados()`), never read from the MOTORISTAS sheet, whose totals depend on whatever period was typed in the PAINEL when the file was saved.
- **Charts**: built lazily per view (`VIEWS[].charts` + `CHART_FN`), because a canvas created while hidden gets stuck at zero size. Series colors follow the validated order `--s1, --s2, --s7, --s3`.
- **Mobile (`max-width:880px`)**: tables become stacked cards via `td::before{content:attr(data-label)}`; every `<td>` must carry `data-label`. The `<meta name="viewport">` tag is load-bearing.
- **Print/PDF**: the PDF button calls `window.print()`; `@media print` lays the current view out on A4 landscape.

### `gerar_dashboard.py` — Excel → JSON/HTML pipeline

Picks the alphabetically-last `.xlsx` in `dados/`. Every column read is declared in a `SPEC_*` list as `(json key, 1-indexed column, expected header, type)`; `validar_cabecalhos()` compares accent-insensitive headers and **aborts listing every mismatch** if the workbook layout changed. Tables that sit mid-sheet are located by header text (`achar_linha()`), not row number.

| JSON key | Sheet | Header row |
|---|---|---|
| `meta` | CAPA (C24 emission, "VERSÃO" row), PARÂMETROS (tolerances, freight base) | — |
| `embarques` | EMBARQUES | 6 |
| `contratos` | CONTRATOS | 6 |
| `posicaoContrato` | POSIÇÃO POR CONTRATO | 6 |
| `transportadoras` / `tarifasFrete` | TRANSPORTADORAS | 6 / found by "CHAVE (TRANSPORTADORA \| LOCAL)" |
| `motoristas` (register only) / `precosCombustivel` | MOTORISTAS | 6 / found by "VIGÊNCIA A PARTIR DE" |
| `clientesLocais` | CLIENTES E LOCAIS | 6 |
| `titulosReceber` | CONTAS A RECEBER | found by "EMBARQUE Nº" (row 21 in v13) |
| `recebimentos`, `notasComplementares`, `pagamentosTransportadora`, `pagamentosMotorista`, `abastecimentosAvulsos` | same-named sheets | 6 |
| `auditoriaResumo` | AUDITORIA | from "VERIFICAÇÃO" under "RESUMO DE OCORRÊNCIAS" through "TOTAL DE OCORRÊNCIAS" |

The per-contract summary at the top of CONTAS A RECEBER and the totals columns of MOTORISTAS are deliberately **not** read: both depend on the PAINEL period.

`conferir_painel()` recomputes 31 indicators for the PAINEL's B7/B8 period and compares them with the PAINEL's column B by row label; on v13 all 31 match (and the accumulated mode matches column C). Treat any divergence as a regression.

`num()` rounds and demotes whole floats to `int`; `txt()` strips whitespace and turns empty strings into `null`.

### Deployment

Public GitHub repo `willian-batista94/terra-biomassa-controle`, GitHub Pages serving `index.html` from `master` → `https://willian-batista94.github.io/terra-biomassa-controle/`. Tracked: `index.html`, `gerar_dashboard.py`, `.gitignore`, `CLAUDE.md`. Everything with data in plain text (`dados/`, `dashboard.html`, `dashboard_slim.json`, `historico_kpis.json`) is gitignored. Note: commits before the encryption change still contain `dashboard_slim.json` in plain text in the git history.

To change the password, just run the generator with the new one; there is no hash stored anywhere.
