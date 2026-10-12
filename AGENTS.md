# AGENTS.md — pravidla pro agenty a přispěvatele

- Nikdy nevolej Save/SaveAs bez výslovného příkazu uživatele.
- Odstranění listu vždy přes potvrzovací dialog s defaultem Ne.
- GUI pouze PyQt6 standardní widgety; žádné vlastní vykreslování.
- COM pouze v `src/excel_backend/com_adapter.py`, používat pouze ve worker vlákně
  (`src/workers/excel_worker.py`, CoInitialize/CoUninitialize). Nesdílet COM objekty mezi vlákny.
- Připojovat se pouze přes GetActiveObject; nevytvářet novou instanci Excelu.
- Chyby překládat přes `core.errors.user_message` — české titulky, žádný obecný „Upozornění".
- Fokus: start na seznamu; po dialozích návrat; po operacích výběr dotčeného/sousedního listu.
- Testy: `pytest -v`. Nemazat ani nezjednodušovat existující testy kvůli průchodu.
- Žádný commit/push bez výslovného požadavku. Žádná vlastní syntéza řeči.
