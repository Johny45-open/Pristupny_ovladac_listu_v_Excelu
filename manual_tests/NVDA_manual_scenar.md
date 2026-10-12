# Manuální testovací scénář — NVDA + Microsoft Excel

Prostředí: Windows 10/11, Excel (otevřený testovací sešit se 3 listy),
NVDA zapnutá, aplikace spuštěná přes `python src\main.py`.
Vše **pouze klávesnicí** (Tab, Shift+Tab, šipky, Enter, Esc).

Příprava sešitu: `Test.xlsx` — List1 (viditelný, aktivní), List2 (viditelný),
Skrytý1 (skrytý). Volitelně Graf1.

## Postup

1. **Start a čtení**
   - [ ] Po startu je fokus v „Seznam listů".
   - [ ] NVDA čte každou položku jako „název, typ, viditelnost, aktivní/neaktivní".
   - [ ] Stav „Aktuální sešit: Test.xlsx" je čitelný.
2. **Tab pořadí**
   - [ ] Tab prochází: seznam → Obnovit → Přejít → Nový → Přejmenovat → Odstranit
         → Vlevo → Vpravo → Na pozici → Skrýt → Zobrazit. Zapište odchylky.
3. **Přechod**
   - [ ] Šipkou vyber List2, Enter. Excel má aktivní List2. Fokus zpět v seznamu na List2.
4. **Nový + přejmenování**
   - [ ] Alt+N → dialog „Vytvořit nový list" → zadej „MujList" → Potvrdit. Nový je vybrán.
   - [ ] Alt+J → dialog „Přejmenovat list…" → změň → Enter (OK) i varianta Esc (beze změny).
5. **Odstranění**
   - [ ] Alt+D → dialog „Odstranit list „MujList"?" obsahuje název, default **Ne**.
   - [ ] Esc/Ne = nic se nesmaže. Ano = smaže se, fokus na sousedovi.
6. **Skrýt / zobrazit**
   - [ ] Vyber List2 → Skrýt. Položka hlásí „skrytý". Zobrazit ji vrátí.
   - [ ] Zobraz velmi skrytý list (připrav v Excelu přes VBA `xlVeryHidden`).
7. **Poslední viditelný**
   - [ ] Skryj vše kromě jednoho → pokus skrýt poslední → dialog „List nelze skrýt". Nic se neskryje.
8. **Přesun**
   - [ ] Alt+V / Alt+B posune a fokus drží přesunutý list. Ověř pořadí v Excelu.
   - [ ] „Přesunout na pozici" → spinbox 1–N → Přesunout.
9. **Externí změna**
   - [ ] V Excelu přejmenuj/přidej list → v aplikaci F5 → seznam odpovídá.
10. **Chybové stavy**
    - [ ] Zavři Excel → Obnovit → „Excel není spuštěn".
    - [ ] Excel bez sešitu → „Žádný aktivní sešit".
    - [ ] Začni editovat buňku (F2) → operace → „Excel nereaguje".
    - [ ] Revize > Chránit sešit → mutace → „Sešit je chráněn". Nic nepadá.
11. **Dialogy**
    - [ ] Každý dialog má smysluplný titulek (ne obecné „Upozornění"), NVDA přečte titulek, text i defaultní tlačítko.

## Záznam

| Krok | Pass/Fail | Poznámka (NVDA verze, Excel verze) |
|---|---|---|
| 1–11 | | |

Nesplněný krok = chyba k opravě, ne „omezení testu".
