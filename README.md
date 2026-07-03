# Report vytíženosti pracovišť (Back Office)

Jupyter notebook, který spočítá vytíženost jednotlivých pracovišť a poboček vůči jejich
denní kapacitě a vygeneruje samostatný interaktivní HTML report.

## Struktura

```
notebooks/
  vytizenost_pracovist_report.ipynb   hlavní notebook — spouštět odshora dolů
  report_lib.py                       výpočty, grafy a sestavení HTML reportu
data/
  sample/                             ukázkový vzorek dat (ze zadání) pro rychlý test
  README.md                           kam nahrát reálná data
output/                               sem se ukládá vygenerovaný HTML report (do gitu se neverzuje)
requirements.txt
```

## Vstupní data

- **`bo_data.xlsx`** — log aktivit: `BRANCH_ID, PRACOVISTE_ID, DATETIME, ZAMESTNANEC, ACTIVITY, DURATION`
  (DURATION v minutách, DATETIME ve formátu `DD.MM.RRRR HH:MM:SS`)
- **`work_spaces.xlsx`** — kapacity poboček: `BRANCH_ID, BRANCH_NAME, NO_WPL, CAPACITY`
  (`NO_WPL` = počet pracovišť na pobočce, `CAPACITY` = kapacita **jednoho** pracoviště v hodinách/den)

Nahrajte oba soubory do `data/` (viz `data/README.md`).

## Spuštění

```bash
pip install -r requirements.txt
jupyter notebook notebooks/vytizenost_pracovist_report.ipynb
```

V notebooku v buňce *1. Konfigurace* nastavte cesty k `bo_data.xlsx` / `work_spaces.xlsx`
(výchozí ukazují na `data/sample/`, aby šlo hned vyzkoušet demo). Pak spusťte
`Kernel → Restart & Run All`. Výsledný samostatný HTML report se uloží do `output/vytizenost_report.html`
a lze ho otevřít v prohlížeči i bez připojení k internetu (Plotly.js je vložený přímo v souboru).

## Co report obsahuje

- **Přehled poboček** — průměrná/max. denní vytíženost vůči kapacitě (`NO_WPL × CAPACITY`), barevně
  odlišeno Nízká / Vysoká / Kritická vytíženost, včetně poboček bez zaznamenaných dat.
- **Trend v čase** — vývoj vytíženosti poboček v jednotlivých dnech (pokud data pokrývají víc dní).
- **Vytíženost jednotlivých pracovišť** — heatmapa den × pracoviště za každou pobočku (rozbalovací sekce).
- **Kapacita: je pracovišť dost?** — porovnává počet skutečně využívaných pracovišť (dle dat) s počtem
  registrovaným v `work_spaces.xlsx`; upozorní, pokud na pobočce reálně přibylo pracovišť.
- **Skladba aktivit a zaměstnanci** — rozložení času dle typu aktivity a nejvytíženější zaměstnanci.
- **Kontrola kolizí rezervací** — zjednodušená detekce časově se překrývajících rezervací téhož pracoviště.
- **Kvalita dat** — přeskočené/neplatné řádky a aktivity patřící neznámé pobočce.

## Konfigurovatelné parametry (`report_lib.py`)

- `BUSINESS_DAYS_ONLY` (výchozí `True`) — vytíženost se počítá jen pro pracovní dny Po–Pá.
- `THRESHOLD_HIGH` / `THRESHOLD_CRITICAL` (výchozí 70 % / 90 %) — hranice pro barevné odlišení stavu.
