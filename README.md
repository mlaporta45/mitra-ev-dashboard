# Mitra EV Financial Dashboard

Real-time financial and operational dashboard for Mitra EV. Upload monthly reports (PDF, Word, PowerPoint, Excel) and the AI automatically extracts KPIs, revenue metrics, and covenant compliance — no manual data entry.

## Features
- Auto-detects reporting period from uploaded documents
- Tracks 14 KPIs + 3 covenants month-over-month
- Interactive trend charts
- Covenant compliance badges (PASS/FAIL)
- Export to formatted Excel workbook
- Public URL — no login required for viewers

## Local Setup

```bash
# 1. Clone / copy the project
cd mitra_ev_dashboard

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add your API key
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edit secrets.toml and paste your ANTHROPIC_API_KEY

# 4. Run
streamlit run app.py
# Opens at http://localhost:8501
```

## Deploy to Streamlit Community Cloud (Free Public URL)

1. Push this folder to a **GitHub repository** (can be private)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **Create app**
3. Select your repo, branch `main`, file `app.py`
4. Click **Advanced settings → Secrets** and paste:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-your-key-here"
   ```
5. Click **Deploy** — you'll get a public URL like `https://mitra-ev.streamlit.app`
6. Share that URL with anyone on your team

## Usage

| Page | What to do |
|------|-----------|
| **Overview** | Select period from dropdown to see KPI cards, covenant badges, revenue chart |
| **Upload** | Drag-drop monthly reports — AI processes and saves automatically |
| **Trends** | Line charts across all periods; covenant threshold lines |
| **Data** | Raw pivot table + Export to Excel button |

## Metrics Tracked
Revenue: EV Vehicle Leases, DCFC Charging Fees, LCFS Credits, Other Revenue, Total Revenue, EBITDA  
Balance Sheet: Cash Position, Total Debt  
Operations: Vehicles in Service, Vehicles Under MLA, Truck MRR, DCFC in Service, DCFC Under SHA, DCFC Avg. Utilization Rate  
Covenants: Minimum Liquidity ($1mm), Tangible Net Worth, DSCR (1.25x)
