# Přístupný ovladač listů Excelu

Samostatná aplikace pro Windows (PyQt6), která ovládá **již spuštěný** Microsoft Excel
a právě aktivní sešit. Plně ovladatelná klávesnicí, optimalizovaná pro čtečku **NVDA**.

## Funkce (MVP)

- Připojení k běžící instanci Excelu, zobrazení názvu aktivního sešitu.
- Přístupný seznam listů: název, pořadí, typ, viditelnost, aktivní/neaktivní.
- Obnovení seznamu (F5 / Alt+O), bezpečné opětovné připojení.
- Přechod na list (Enter), nový list, přejmenování, odstranění s potvrzením (default **Ne**).
- Přesun vlevo/vpravo + přesun na konkrétní pozici.
- Skrytí / zobrazení (včetně velmi skrytých `xlVeryHidden`).
- Ochrana: zákaz skrýt poslední viditelný list, srozumitelné chyby při chráněné struktuře.
- Aplikace **nikdy** automaticky neukládá sešit (žádné Save/SaveAs).

## Instalace

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Požadavky: Windows 10/11, Python 3.11+, Microsoft Excel, NVDA (doporučeno).

## Spuštění

```bat
python src\main.py
```

## Testy

```bat
pytest -v
```

Automatické testy používají `FakeExcelAdapter` a nevyžadují Excel.

## Ovládání klávesnicí

| Akce | Klávesa |
|---|---|
| Pohyb v seznamu | Šipky nahoru/dolů |
| Přejít na list | Enter / Alt+P |
| Obnovit | F5 / Alt+O |
| Nový list | Alt+N |
| Přejmenovat | Alt+J |
| Odstranit | Alt+D (potvrzení, default Ne) |
| Vlevo / vpravo | Alt+V / Alt+B |
| Na pozici | Alt+N? ne — tlačítko „Přesunout na pozici" Alt+A? viz UI: Alt+N je nový; pozice přes tlačítko |
| Skrýt / Zobrazit | Alt+S / Alt+Z |
| Zavřít dialog bez změny | Esc |

Přesné popisky tlačítek viz `src/ui/main_window.py`. Fokus startuje v seznamu listů.

## Limity

- Změny struktury listů jsou v Excelu okamžité i bez uložení. Uložení proveďte v Excelu (Ctrl+S).
  Aplikace záměrně neukládá.
- Pokud Excel edituje buňku nebo má otevřený modální dialog, operace hlásí „Excel nereaguje".
- Při chráněné struktuře (Revize > Chránit sešit) jsou mutace zakázány.
- Hlasový výstup zajišťuje výhradně čtečka obrazovky; aplikace nemá vlastní syntézu.

## Struktura

Viz `AGENTS.md` a implementační plán. Vrstvy: `ui/` → `workers/excel_worker.py` (QThread+COM)
→ `controller/sheet_controller.py` → `excel_backend/com_adapter.py`.
