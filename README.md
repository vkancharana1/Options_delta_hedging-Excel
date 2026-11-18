# Options Delta Analysis & Dynamic Delta Hedging in Excel

**An Excel-based derivatives risk-management project demonstrating Black–Scholes delta estimation, option-chain analysis, and the mechanics of rebalancing a delta hedge over time.**

**Tools:** Microsoft Excel · Black–Scholes framework · `LN` · `SQRT` · `NORM.DIST` · `ROUND` · absolute/relative cell references · cash-flow and interest calculations

## Project overview

Option prices respond to movements in the underlying asset, volatility, and time to expiration. **Delta** estimates how much an option's value changes for a small change in the underlying price. Because delta changes over time, hedging an options position requires periodically adjusting the amount of the underlying asset held.

This repository contains two complementary Excel exercises:

1. **Delta Calculation.xlsx** — Explore a NIFTY options-chain snapshot and calculate illustrative Black–Scholes call deltas for selected strikes.
2. **Dynamic Delta Hedging.xlsx** — Follow a short-call position through a 20-week hedging scenario, tracking delta, stock purchases/sales, financing interest, and the terminal hedge cost.

> **Scope:** These are spreadsheet-based educational models, not a live trading system or an investment recommendation. The NIFTY option-chain analysis and the stock-based hedge simulation are separate exercises; they do not represent one continuous backtest on the same underlying.

## Key figures

| Metric | Workbook | Value |
|---|---|---:|
| NIFTY strike-price observations | Delta Calculation | **78** |
| Strike-price range | Delta Calculation | **23,950–27,800** |
| Rows with additional delta analysis populated | Delta Calculation | **12** (selected strikes) |
| Short call options modeled | Dynamic Delta Hedging | **100,000** |
| Starting stock price / strike price | Dynamic Delta Hedging | **$49 / $50** |
| Assumed annual volatility / interest rate | Dynamic Delta Hedging | **20% / 5%** |
| Hedging horizon | Dynamic Delta Hedging | **20 weeks** |
| Weekly intervals | Dynamic Delta Hedging | **20** (week 20 to expiry) |
| Initial delta / stock hedge | Dynamic Delta Hedging | **0.5216 / 52,200 shares** |
| Final stock price at expiry | Dynamic Delta Hedging | **$57.25** |
| Modeled terminal hedge cost | Dynamic Delta Hedging | **$263,305.47** |

The dollar sign above refers to the units used in the **hedging simulation**, not the NIFTY option-chain data.

## Repository files

```text
.
├── README.md
├── Delta Calculation.xlsx
└── Dynamic Delta Hedging.xlsx
```

## 1. NIFTY options-chain delta analysis

**File:** [`Delta Calculation.xlsx`](./Delta%20Calculation.xlsx)

The workbook includes an option-chain snapshot labeled **2025-11-18**, with **78 strike rows** and fields for call and put prices, implied volatility (IV), delta, gamma, theta, vega, volume, open interest (OI), and bid/offer prices.

The supplementary calculations demonstrate how the *call delta* can be estimated using the Black–Scholes framework for selected strikes. The calculation uses the underlying price, strike price, implied volatility, risk-free rate, and remaining time to expiry.

For a non-dividend-paying underlying, the standard call-delta expression is:

\[
\Delta_{call}=N(d_1), \qquad
 d_1=\frac{\ln(S/K)+(r+\sigma^2/2)T}{\sigma\sqrt{T}}
\]

where:

- **S** = underlying price
- **K** = option strike price
- **r** = annual risk-free interest rate
- **σ** = annualized volatility
- **T** = time to expiry, in years
- **N(·)** = standard normal cumulative distribution function

**Workbook-specific note:** The model includes a reference underlying value of **25,910.05**, a **4-day** expiry assumption, and an **8%** rate in the illustrative formulas. Additional delta-related values are populated across **12 selected strike rows**, but the cells are **not all backed by formulas** in the uploaded workbook. The option chain's vendor-provided Greeks should not be mistaken for independently computed outputs. The formulas also vary in implementation between some example rows, so treat the sheet as a demonstration rather than a fully validated, standardized pricing engine.

## 2. Dynamic delta hedging simulation

**File:** [`Dynamic Delta Hedging.xlsx`](./Dynamic%20Delta%20Hedging.xlsx)

The simulation starts from a **short position in 100,000 call options**. Selling calls creates negative call-delta exposure; buying the underlying stock offsets that exposure. As stock prices and time to expiry change, the target stock position is updated.

### Inputs

| Input | Value |
|---|---:|
| Call options sold | 100,000 |
| Initial stock price | $49.00 |
| Strike price | $50.00 |
| Annual volatility | 20% |
| Annual risk-free interest rate | 5% |
| Initial time to expiry | 20 / 52 years |
| Rebalancing frequency | Weekly |

### How the hedge works

1. **Estimate delta:** Compute \(d_1\) from the Black–Scholes inputs, then obtain call delta with `NORM.DIST`.
2. **Set the hedge position:** Multiply call delta by **100,000 options** and round shares to the nearest **100** in the spreadsheet.
3. **Rebalance each week:** The change in shares held equals the new target hedge minus the previous stock position. A positive value indicates purchases; a negative value indicates sales.
4. **Account for cash and financing:** Track stock-purchase/sale cash flows, cumulative cost, and interest at the assumed annual rate.
5. **Evaluate the terminal outcome:** Compare final cumulative hedge cost against the strike-price obligation for the 100,000 short calls.

### Representative Excel formulas

The following formulas are used in the hedging workbook (first scenario row):

```excel
' Time to expiry in years
=F2/52

' Black–Scholes d1
=(LN(H2/$C$5)+($C$4+0.5*$C$3*$C$3)*G2)/($C$3*SQRT(G2))

' Call delta
=NORM.DIST(I2,0,1,TRUE)

' Target shares, rounded to the nearest 100
=ROUND(J2*$C$2,-2)

' Shares purchased/sold in the next period
=K3-K2

' Financing cost for one week
=N2*$C$4/52

' Cumulative financing-adjusted stock cost
=N2+M3+O2

' Terminal reported hedge cost
=N22*1000-C5*C2
```

> **Important:** The shares-cost column is displayed **in thousands**. The workbook's final formula scales cumulative cost back up by 1,000 before subtracting the strike obligation.

### Illustrative results

| Stage | Stock price | Call delta | Shares held |
|---|---:|---:|---:|
| Week 20 (start) | $49.00 | 0.5216 | 52,200 |
| Week 10 | $49.88 | 0.5502 | 55,000 |
| Week 5 | $51.88 | 0.7592 | 75,900 |
| Week 1 | $55.87 | ~1.0000 | 100,000 |
| Week 0 (expiry) | $57.25 | — | — |

At expiry, the workbook reports a financing-adjusted cumulative stock cost of approximately **$5,263,305**, against a **$5,000,000** strike-price obligation, resulting in a **$263,305.47 modeled hedging cost**.

**Interpretation:** The result illustrates how a hedge generates trading and financing cash flows as delta evolves. It is **not** evidence of $263K of profit, model accuracy, hedge effectiveness relative to a benchmark, or actual realized trading performance. The workbook does not establish a comparison with an unhedged strategy, bid–ask/commission costs, or real executions.

| Spreadsheet capability | Where it appears | Practical purpose |
|---|---|---|
| Formula-driven financial modeling | Dynamic Delta Hedging | Connect option sensitivities to hedge positions and cash flows |
| `LN`, `SQRT`, `NORM.DIST` | Both workbooks | Implement Black–Scholes-style calculations |
| Absolute/relative references (`$C$5`, `H2`) | Both workbooks | Separate fixed assumptions from row-level market inputs |
| `ROUND` and position-difference calculations | Dynamic Delta Hedging | Set share targets and calculate rebalancing trades |
| Rolling cumulative-cost and interest calculations | Dynamic Delta Hedging | Model financing and terminal cost |
| Structured market-data review | Delta Calculation | Examine strike-level IV, Greeks, prices, volume, and OI |

**Not claimed:** VBA, macros, Power Query, PivotTables, automated external data refresh, or live execution are not evidenced by these files.

## How to explore the workbooks

1. Download or clone this repository.
2. Open the `.xlsx` files in **Microsoft Excel** (desktop Excel is recommended).
3. Start with **Delta Calculation.xlsx** to see the market-data inputs and selected theoretical delta calculations.
4. Open **Dynamic Delta Hedging.xlsx** and inspect the assumptions in **C2:C5**, the weekly simulation in **F:O**, and the final result in **H24**.
5. Change an assumption in a **copy** of the hedging file and check which outputs recalculate. **Note:** Not every displayed historical/simulation value is formula-linked in the supplied workbook; some are stored values. Changing one input does not guarantee the full scenario recalculates automatically.

## Limitations and potential extensions

- **Illustrative inputs:** The hedging scenario uses a specified weekly price path rather than a documented live dataset.
- **Incomplete formula propagation:** Many weekly output cells contain stored numbers, so this is not yet a fully parameterized, one-click scenario engine.
- **Model assumptions:** The simple Black–Scholes framework assumes constant volatility/rates and idealized trading assumptions; actual markets have additional frictions.
- **No formal backtest or benchmarking:** The reported hedge cost is a scenario outcome, not a measured improvement versus a benchmark.
- **Future work:** Extend formulas consistently across all rows, validate Greeks against reference calculations, add line charts and scenario stress tests, incorporate transaction costs, and compare hedged versus unhedged P&L.

---

*Educational quantitative finance project. Models and market-data fields should be independently checked before any real-world use.*
