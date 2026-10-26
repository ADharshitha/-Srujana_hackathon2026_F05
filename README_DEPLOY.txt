FinTech Tester + Hypothetical Tracker

Files:
- app.py: Python HTTP server and Base44 API integration
- index.html: existing dashboard, debt alerts, monthly Excel report, with navigation link
- hypothetical.html: new "What If I Had Spent Less?" scenario page
- requirements.txt: Python dependency
- render.yaml: optional Render Blueprint configuration

Run locally (Windows PowerShell):
1. Put all files in the same folder.
2. Install Python 3.10+.
3. Run: py -m pip install -r requirements.txt
4. Set your Base44 token for this terminal:
   $env:BASE44_TOKEN = "YOUR_PRIVATE_BASE44_TOKEN"
5. Run: py app.py
6. Open http://localhost:10000

Render:
1. Push these files to your GitHub repository.
2. In Render, create a Web Service from that repository.
3. Build command: pip install -r requirements.txt
4. Start command: python app.py
5. Add BASE44_TOKEN under Environment Variables as a secret.
6. Deploy. Do not commit the real token to GitHub.

Notes:
- The hypothetical tracker uses the existing /api/dashboard endpoint; no new API endpoint is needed.
- Recurring payments are heuristic candidates (same merchant/category in at least two distinct months), not guaranteed subscriptions.
- Budget utilisation uses the budget entered on the hypothetical page.
- For transactions with no explicit direction/type, only negative amounts are treated as expenses. If your Base44 schema stores all amounts as positive and uses a different direction field, adjust isExpense() in hypothetical.html to match your schema.
- On Render, the local debts.json file may not persist across redeploys/restarts unless you use persistent storage or a database.
