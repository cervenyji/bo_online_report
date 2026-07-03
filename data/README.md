# Vstupní data

Do této složky nahrajte dva soubory (přesně pod těmito názvy, nebo si cestu upravte
v notebooku v buňce *Konfigurace*):

- `bo_data.xlsx` — log aktivit zaměstnanců na pracovištích:
  `BRANCH_ID, PRACOVISTE_ID, DATETIME, ZAMESTNANEC, ACTIVITY, DURATION` (DURATION v minutách)
- `work_spaces.xlsx` — kapacity poboček:
  `BRANCH_ID, BRANCH_NAME, NO_WPL, CAPACITY` (NO_WPL = počet pracovišť, CAPACITY = kapacita
  jednoho pracoviště v hodinách/den)

Oba soubory jsou v `.gitignore`, takže se s reálnými daty neverzují do gitu.

## sample/

Obsahuje malý ukázkový vzorek (přesně data ze zadání) pro rychlé otestování notebooku
bez nutnosti mít po ruce ostrá data.
