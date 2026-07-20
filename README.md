# Personal Investment Dashboard

A multi-currency, multi-market investment tracker built with **Excel** (data model) and **Power BI** (dashboards). Designed to answer the questions that brokerage apps don't: What is my total net worth today? How much passive income am I generating? And are my stock losses really as bad as they look once you account for dividends and fixed income?

---

## What This Project Does

Most investors hold assets across multiple platforms, currencies, and instrument types — and most apps only show you one piece at a time. This dashboard consolidates everything into five pages:

| Page | Question it answers |
|---|---|
| **Patrimonio** | What is my total net worth across all asset types? |
| **Portafolio** | How are my stock positions performing, including closed trades? |
| **Dividendos** | How much dividend income am I receiving, and how much goes to taxes? |
| **Renta Fija** | How much am I earning from fixed-income instruments, and how has that changed over time? |
| **Diversificación** | Am I too concentrated in one sector, region, or risk level? |

---

## Screenshots

### Patrimonio — Total Net Worth
![Patrimonio](screenshots/Patrimonio.png)

### Portafolio — Stock Performance
![Portafolio](screenshots/portafolio.png)

### Dividendos — Dividend Income
![Dividendos](screenshots/dividendos.png)

### Renta Fija — Fixed Income
![Renta Fija](screenshots/rentafija.png)

### Diversificación — Portfolio Composition
![Diversificacion](screenshots/diversificacion.png)

### Data Model
![Data Model](screenshots/data_model.png)

---

## Data Model

The Excel workbook (`Investments_Demo.xlsx`) contains these tables:

```
Fact_Transactions      — Every buy and sell, with price, FX rate, commissions, taxes
Fact_Dividends         — Dividend income with gross amount, withholding tax, net received
Fact_FixedIncome       — Fixed-rate instruments with start/end dates, principal, annual rate
Fact_CurrentPrices     — Current market price for each active position
Fact_PortfolioSnapshot — Monthly snapshot of total portfolio value by asset group
Fact_FXTransactions    — Currency purchases (USD, EUR, etc.)
Dim_Assets             — Asset metadata: sector, region, risk level, dividend frequency
Dim_FXRates            — Historical exchange rates by month
```

Each transaction stores the FX rate at the time of the operation, so the cost basis in MXN is always historically accurate — not recalculated at today's rate.

---

## Key DAX Patterns

### Avoiding filter context errors between tables
When calculating the current value of a position, you can't compare columns from two different tables inside a `FILTER`. The fix is to use `VAR` to capture the asset name as a scalar first:

```dax
Valor Actual MXN =
SUMX(
    VALUES(Fact_Transactions[AssetName]),
    VAR activo = Fact_Transactions[AssetName]
    VAR shares = CALCULATE(
        SUM(Fact_Transactions[Quantity]),
        Fact_Transactions[TransactionType] = "Buy",
        Fact_Transactions[Status] = "Active"
    )
    VAR precio = CALCULATE(
        MAX(Fact_CurrentPrices[CurrentPrice_MXN]),
        Fact_CurrentPrices[AssetName] = activo,
        Fact_CurrentPrices[Status] = "Active"
    )
    RETURN shares * precio
)
```

### Historical fixed income — ALLSELECTED vs ALL
To show estimated monthly income across all historical periods while still respecting slicer selections:

```dax
Ingreso Mensual por Fecha =
SUMX(
    FILTER(
        ALLSELECTED(Fact_FixedIncome),
        Fact_FixedIncome[StartDate] <= MAX('Calendar'[Date]) &&
        Fact_FixedIncome[EndDate] >= MIN('Calendar'[Date])
    ),
    Fact_FixedIncome[Principal_MXN] * Fact_FixedIncome[AnnualRate] / 12
)
```

`ALLSELECTED` removes automatic context filters but keeps slicer selections — so the chart shows the full date range while still responding to the instrument filter.

### Separating realized vs unrealized P&L
Active positions use current market price. Closed positions use the actual sale price:

```dax
Valor Actual MXN =
SUMX(
    VALUES(Fact_Transactions[AssetName]),
    VAR activo = Fact_Transactions[AssetName]
    VAR esClosed =
        CALCULATE(
            COUNTROWS(Fact_Transactions),
            Fact_Transactions[TransactionType] = "Sell",
            Fact_Transactions[Status] = "Closed"
        ) > 0
    VAR valorVenta =
        CALCULATE(
            SUM(Fact_Transactions[NetAmount_MXN]),
            Fact_Transactions[TransactionType] = "Sell",
            Fact_Transactions[Status] = "Closed"
        )
    VAR shares = CALCULATE(SUM(Fact_Transactions[Quantity]),
        Fact_Transactions[TransactionType] = "Buy")
    VAR precio = CALCULATE(
        MAX(Fact_CurrentPrices[CurrentPrice_MXN]),
        Fact_CurrentPrices[AssetName] = activo,
        Fact_CurrentPrices[Status] = "Active"
    )
    RETURN IF(esClosed, valorVenta, shares * precio)
)
```

---

## Multi-Currency Design

This dashboard tracks assets in MXN, USD, EUR, and GBP. The approach:

- Every transaction stores the FX rate at the time of purchase (`FXRate_to_MXN`)
- Cost basis in MXN reflects historical rates — not today's
- The Patrimonio page shows totals in MXN, USD, and EUR using the most recent rate from `Fact_PortfolioSnapshot`
- Some ETFs trade on both the Mexican market (MXN) and US market (USD) as separate positions — handled with an `AssetKey` column that uniquely identifies each combination (e.g., `VOO_MX` vs `VOO_US`)

---

## Tax Handling

The model tracks two types of withholding:

- **Mexican stocks (ISR)**: 10% retained on gross dividend
- **Foreign ETFs via SIC (W-8BEN)**: 10% US withholding on dividends (reduced from 30% with a W-8BEN form on file)
- **FIBRAs** (Mexican REITs): effective rate varies per distribution — must be derived from actual broker figures, as each quarterly payment is split across taxable income, return of capital, etc.

Withholding is stored at the transaction level in `Fact_Dividends`, so the dashboard always shows both gross and net dividend income.

---

## How to Use This

### Requirements
- Microsoft Excel (to edit the data)
- [Power BI Desktop](https://powerbi.microsoft.com/desktop/) (free)

### Setup
1. Download `Investments_Demo.xlsx` and the `.pbix` file
2. Open the `.pbix` in Power BI Desktop
3. Go to **Home → Transform data → Data source settings**
4. Update the file path to point to your local copy of the Excel file
5. Click **Refresh** — the dashboard will load with the demo data

### Adapting to your own data
Replace the demo data in each Excel table with your own transactions. The key things to match:
- Use the same column names (Power BI references them by name)
- Keep the `AssetKey` format consistent (`TICKER_MARKET`, e.g. `AAPL_US`)
- Update `Dim_Assets` with your actual holdings and their metadata
- Update `Fact_CurrentPrices` regularly with current market prices

---

## Project Structure

```
├── Investments_Demo.xlsx     # Demo data (fictional, same structure as real workbook)
├── Dashboard.pbix            # Power BI file connected to the demo Excel
├── backgrounds/              # PNG background images for each dashboard page (1280×720)
│   ├── bg_patrimonio.png
│   ├── bg_portafolio.png
│   ├── bg_dividendos.png
│   ├── bg_rentafija.png
│   └── bg_diversificacion.png
├── screenshots/              # Dashboard screenshots for this README
└── README.md
```

---

## Related Article

This project is documented in detail on Medium:
**[How I Built a Personal Investment Dashboard with Excel, Power BI, and a Lot of DAX Headaches](#)**
*(link to be added after publication)*

---

## Notes

- The demo file uses fictional tickers, amounts, and dates — the structure is identical to the real workbook but no real financial data is included
- Fixed income instruments are modeled with explicit start/end dates and annual rates — digital savings accounts (daily interest, no fixed term) are handled by updating the balance monthly and marking the most recent record as Active
- The Calendar table is generated dynamically from the minimum date across all fact tables, so it automatically extends as you add historical data

---

*Built by [your name] · Feel free to open an issue if you have questions about the data model or DAX measures*
