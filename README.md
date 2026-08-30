# 📊 Personal Finance Dashboard — Streamlit App

A dark-teal glassmorphic Streamlit dashboard that connects directly to your
`Personal_Finance_Dashboard - Year - Month.xlsx` workbook (all 10 sheets),
lets you edit every table like Excel, and recalculates every KPI and chart
in real time.

## 1. Install (one time)

You need Python 3.9+. In a terminal:

```bash
cd pfd_app
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Run

```bash
streamlit run app.py
```

Your browser opens automatically at `http://localhost:8501`.

## 3. Load your data

- Use the **📂 Upload Excel File** control in the sidebar to upload your
  `.xlsx` file, **or**
- Place a copy of your workbook in the same folder as `app.py`, named
  `Personal_Finance_Dashboard_-_Year_-_Month.xlsx` — the app auto-loads it
  on startup (a sample copy is already included here so you can try the app
  immediately).

## 4. What each tab does

| Tab | Sheet(s) it manages | Notes |
|---|---|---|
| 📊 Dashboard Overview | 01, 03, 07 (read/derived) | KPI cards, donut charts, gauge, funnel, cash-flow trend, debt progress — all recomputed live from your transactions |
| 💵 Income & Expense | 02 | Editable transaction log (`st.data_editor`, add/delete rows) |
| 💳 Debt & Assets | 05, 06 | Editable debt & asset tables; Remaining/Status are auto-computed |
| 🛟 Emergency & Savings | 04, 08 | EF target inputs + gauge; editable saving/investment allocation log |
| 📅 Monthly Report | 09 (derived) | 12-month table + trend charts for the selected year |
| ⚙️ Settings | 10 | Editable dropdown master lists |

Use the **ປີ (Year)** / **ເດືອນ (Month)** selectors in the sidebar to filter
everything at once.

## 5. Saving your changes

Click **💾 ບັນທຶກ/ດາວໂຫລດ Excel** in the sidebar at any time. This writes all
your edits back into a copy of your *original* workbook — formulas on sheets
like `01_Dashboard`, `03_Cash_Flow`, `07_Net_Worth`, and `09_Monthly_Report`
are left intact (not overwritten with static numbers), so they recalculate
normally the next time you open the file in Excel.

> ⚠️ If you add far more rows than the original sheet had (e.g. hundreds
> more transactions), Excel's named ranges are automatically extended to
> match — but double-check the `01_Dashboard` and `07_Net_Worth` sheets
> after opening, since a couple of the Dashboard's KPI cells use Excel
> array formulas that this app does not attempt to rewrite.

## 6. How the numbers are calculated

The app mirrors the workbook's own named-range formulas 1:1 in pandas
(`SUMIFS`-equivalent filters on Type / Category / Need_Want / Month / Year),
so the figures you see in the app match what Excel calculates from the same
data. This was verified against the sample workbook's cached formula results
during development.

## 7. Currency & language

Amounts are formatted as Lao Kip (`₭`). Column headers keep the original
Lao/English bilingual labels from your workbook so the app stays consistent
with the source file.
