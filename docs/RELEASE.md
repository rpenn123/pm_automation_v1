# Release checklist

1. Deploy to **TEST** (Update.bat)
2. Open TEST Sheet → refresh → **Project Actions → Run Full Setup** if scopes/triggers changed
3. Run smoke test:
   - Progress sync Forecasting ↔ Upcoming
   - Permits=approved → Upcoming
   - Delivered=TRUE → Inventory_Elevators
   - Progress=In Progress → Framing
4. Check **Executions** for green runs
5. Deploy to **PROD** (Update-Prod.bat)
6. Repeat Full Setup if scopes/triggers changed
7. Re-check Executions after first day
