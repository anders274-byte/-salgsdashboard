# Salgsdashbord — BAMA Blomster

Selvstendig HTML-dashbord (`index.html`) med totalsalg hittil i år, salg per
uke/måned/år og volumoversikt med leveringsgrad (LG) og direktesalg.
Siden er helt frittstående (Chart.js er bakt inn) og kan åpnes rett fra disk
eller deles som én fil.

## Regenerere med ferske data

```bash
pip install openpyxl
python3 generator.py Salg_konvertert_output.xlsm FCA_Rapport.xlsx [--dato YYYY-MM-DD]
```

- `--dato` overstyrer «i dag» (standard: dagens dato). Brukes til YTD-kutt og markører.
- Output skrives til `index.html` i samme mappe.

## Filer

| Fil | Beskrivelse |
| --- | --- |
| `index.html` | Generert dashbord (data + Chart.js innebygd) |
| `generator.py` | Leser kildefilene, aggregerer og fyller malen |
| `template.html` | HTML/CSS/JS-mal med plassholdere `__DATA__` og `/*__CHARTJS__*/` |
| `vendor/chart.umd.js` | Chart.js v4.4.9 (MIT), bakes inn i siden |
| `.claude/skills/run-salgsdashboard/` | Skill + headless smoke-driver (se SKILL.md) |

## Definisjoner

- **Salgsvolum** = «Antall» (kolli/D-pak) fra `Salg_konvertert`-arket. Rader
  merket «Rullerende» er fremtidig ordreinngang og vises kun som prognoselinje.
- **LG (leveringsgrad)** = levert ÷ bestilt.
- Salgsfilen omfatter kun **Tranby-volum (eks. direkteleveranser)**. Dette er
  verifisert mot «Daglig stand-up»-arket i FCA-rapporten: dashbordets totaler
  minus siste dag treffer stand-up-tallene eksakt for både 2025 og 2026.
- **Direktesalg** (PO til kundens DC) hentes fra raden «Direktelev YTD …
  (info)» i stand-up-arket. Tallet finnes kun som hittil-i-år-verdi i kildene,
  ikke per uke/måned, og kommer i tillegg til Tranby-volumet.
- **Uker** er ISO-uker (merk: Uke/Måned/År-arkene i salgsfilen blander
  kalenderår og ISO-uke rundt årsskiftet og kan avvike noe).
