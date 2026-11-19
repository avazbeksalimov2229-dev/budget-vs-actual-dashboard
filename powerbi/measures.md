# DAX: Project Budget vs Actual Dashboard

All measures live in an empty table called **`_Measures`**. The folder names are the **Display folder** values.

**How "as of" works.** The report has one *As of month* slicer. Every "to date" measure takes the end of the
selected month (capped at the last month with actuals) and adds up everything up to that day. On a chart with
months on the axis, each month gets its own as-of date, so the same measures draw cumulative S-curves.

---

## 1. Calculated tables

```dax
dim_date =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2024, 7, 1 ), DATE ( 2027, 12, 31 ) ),
    "Year", YEAR ( [Date] ),
    "Quarter", "Q" & QUARTER ( [Date] ),
    "Year Quarter", YEAR ( [Date] ) & "-Q" & QUARTER ( [Date] ),
    "Month Number", MONTH ( [Date] ),
    "Month", FORMAT ( [Date], "MMM" ),
    "Year Month", FORMAT ( [Date], "YYYY-MM" ),
    "Month Start", DATE ( YEAR ( [Date] ), MONTH ( [Date] ), 1 ),
    "Month End", EOMONTH ( [Date], 0 ),
    -- months that already have actuals: used to filter the as-of slicer and the trend charts
    "Is Closed Month", EOMONTH ( [Date], 0 ) <= EOMONTH ( MAX ( fact_actuals[posting_date] ), 0 )
)
```
Mark as date table. Sort `Month` by `Month Number`.

```dax
-- Field parameter: lets the viewer switch the chart axis
View By =
{
    ( "Project", NAMEOF ( dim_project[project_id] ), 0 ),
    ( "Project manager", NAMEOF ( dim_pm[pm_name] ), 1 ),
    ( "Project type", NAMEOF ( dim_project[project_type] ), 2 ),
    ( "Cost category", NAMEOF ( dim_cost_code[cost_category] ), 3 )
}
```

## 2. Relationships (all one-to-many, single direction)

| From (one) | To (many) |
|---|---|
| `dim_pm[pm_id]` | `dim_project[pm_id]` |
| `dim_project[project_id]` | every fact table |
| `dim_cost_code[cost_code]` | `fact_budget`, `fact_actuals`, `fact_commitments`, `fact_change_orders`, `fact_forecast` |
| `dim_date[Date]` | `fact_budget[month_start]`, `fact_actuals[posting_date]`, `fact_commitments[commitment_date]`, `fact_change_orders[approved_date]`, `fact_progress[week_ending]`, `fact_forecast[forecast_month]`, `fact_billing[billing_period_end]` |

## 3. Row-level security

| Role | Table filter on `dim_pm` | Who |
|---|---|---|
| **Project Manager** | `[pm_email] = USERPRINCIPALNAME ()` | each PM sees only their own projects |
| **Executive** | none | sees the whole portfolio |

The filter flows `dim_pm → dim_project → every fact table`, so one rule secures the whole model.

---

## 4. As-of date (folder: `0 As of`)

```dax
Data Through =
EOMONTH ( CALCULATE ( MAX ( fact_actuals[posting_date] ), REMOVEFILTERS () ), 0 )

As Of Date = MIN ( MAX ( dim_date[Date] ), [Data Through] )

-- TRUE when the whole period on a chart axis is after the last actuals
Is Future Period = MIN ( dim_date[Date] ) > [Data Through]

As Of Label = "Data as of " & FORMAT ( [As Of Date], "d mmm yyyy" )
```

## 5. Budget and cost (folder: `1 Cost`)

```dax
Original Budget =
CALCULATE ( SUM ( fact_budget[amount] ), fact_budget[budget_type] = "Original", REMOVEFILTERS ( dim_date ) )

-- owner change orders approved by the as-of date
Approved CO Cost =
VAR d = [As Of Date]
RETURN
    CALCULATE (
        SUM ( fact_change_orders[cost_amount] ),
        fact_change_orders[status] = "Approved",
        REMOVEFILTERS ( dim_date ),
        fact_change_orders[approved_date] <= d
    ) + 0

Revised Budget = [Original Budget] + [Approved CO Cost]

Planned Cost (month) = SUM ( fact_budget[amount] )

-- cumulative phased budget: the planned S-curve keeps going after the as-of date, up to the planned finish
Planned Value =
VAR d = MAX ( dim_date[Date] )
VAR lastPlanMonth = EOMONTH ( CALCULATE ( MAX ( fact_budget[month_start] ), REMOVEFILTERS ( dim_date ) ), 0 )
RETURN
    IF (
        MIN ( dim_date[Date] ) <= lastPlanMonth,
        CALCULATE ( SUM ( fact_budget[amount] ), REMOVEFILTERS ( dim_date ), fact_budget[month_start] <= d )
    )

Actual Cost (month) = IF ( NOT [Is Future Period], SUM ( fact_actuals[amount] ) )

Actual Cost =
VAR d = [As Of Date]
RETURN
    IF (
        NOT [Is Future Period],
        CALCULATE ( SUM ( fact_actuals[amount] ), REMOVEFILTERS ( dim_date ), fact_actuals[posting_date] <= d )
    )

Committed Cost =
VAR d = [As Of Date]
RETURN
    CALCULATE ( SUM ( fact_commitments[amount] ), REMOVEFILTERS ( dim_date ), fact_commitments[commitment_date] <= d )

Invoiced on Commitments =
VAR d = [As Of Date]
RETURN
    CALCULATE (
        SUM ( fact_actuals[amount] ),
        REMOVEFILTERS ( dim_date ),
        fact_actuals[posting_date] <= d,
        NOT ISBLANK ( fact_actuals[po_id] )
    )

-- what is still owed on signed subcontracts and POs, never below zero per project and cost code
Open Commitments =
SUMX (
    CROSSJOIN ( VALUES ( dim_project[project_id] ), VALUES ( dim_cost_code[cost_code] ) ),
    MAX ( 0, [Committed Cost] - [Invoiced on Commitments] )
)

-- early warning: signed commitments already above the budget of their cost code
Committed Over Budget =
SUMX (
    CROSSJOIN ( VALUES ( dim_project[project_id] ), VALUES ( dim_cost_code[cost_code] ) ),
    VAR over = [Committed Cost] - [Revised Budget]
    RETURN IF ( over > 0, over )
)

Committed % of Budget = DIVIDE ( [Committed Cost], [Revised Budget] )

-- costs coded to 99-999: they need recoding before month-end close
Unallocated Cost = CALCULATE ( [Actual Cost], dim_cost_code[cost_code] = "99-999" )
```

## 6. Earned value (folder: `2 Earned value`)

Progress is reported per project, so earned value is calculated per project and then added up.

```dax
-- physical % complete from the PM's latest weekly entry on or before the as-of date
Earned Value =
VAR d = [As Of Date]
RETURN
    IF (
        NOT [Is Future Period],
        SUMX (
            VALUES ( dim_project[project_id] ),
            VAR lastWeek =
                CALCULATE ( MAX ( fact_progress[week_ending] ), REMOVEFILTERS ( dim_date ), fact_progress[week_ending] <= d )
            VAR pct =
                CALCULATE ( MAX ( fact_progress[pct_complete] ), REMOVEFILTERS ( dim_date ), fact_progress[week_ending] = lastWeek )
            RETURN pct * [Revised Budget]
        )
    )

% Complete = DIVIDE ( [Earned Value], [Revised Budget] )

CPI = DIVIDE ( [Earned Value], [Actual Cost] )

SPI =
VAR d = [As Of Date]
VAR pv = CALCULATE ( SUM ( fact_budget[amount] ), REMOVEFILTERS ( dim_date ), fact_budget[month_start] <= d )
RETURN DIVIDE ( [Earned Value], pv )

Cost Variance (EV - AC) = [Earned Value] - [Actual Cost]
```

## 7. Forecast (folder: `3 Forecast`)

```dax
-- PM forecast to complete: the latest monthly snapshot on or before the as-of date, per project
Forecast to Complete =
VAR d = [As Of Date]
RETURN
    SUMX (
        VALUES ( dim_project[project_id] ),
        VAR snap =
            CALCULATE (
                MAX ( fact_forecast[forecast_month] ),
                REMOVEFILTERS ( dim_date ),
                REMOVEFILTERS ( dim_cost_code ),
                fact_forecast[forecast_month] <= d
            )
        RETURN
            CALCULATE ( SUM ( fact_forecast[forecast_to_complete] ), REMOVEFILTERS ( dim_date ), fact_forecast[forecast_month] = snap )
    )

-- Estimate at completion, as forecast by the project manager
EAC = IF ( NOT [Is Future Period], [Actual Cost] + [Forecast to Complete] )

-- Independent check: remaining work at the cost efficiency achieved so far.
-- CPI is only stable after about 20% complete, so younger projects keep the PM forecast.
EAC (CPI Trend) =
SUMX (
    VALUES ( dim_project[project_id] ),
    VAR ac = [Actual Cost]
    VAR ev = [Earned Value]
    VAR bac = [Revised Budget]
    VAR cpi = DIVIDE ( ev, ac )
    RETURN IF ( ISBLANK ( cpi ) || DIVIDE ( ev, bac ) < 0.2, [EAC], ac + DIVIDE ( bac - ev, cpi ) )
)

-- positive = under budget, negative = over budget
Variance at Completion = [Revised Budget] - [EAC]

Variance at Completion % = DIVIDE ( [Variance at Completion], [Revised Budget] )

-- for the waterfall: how much each part adds to (or saves from) the budget
Overrun vs Budget = [EAC] - [Revised Budget]

Overrun (CPI Trend) = [EAC (CPI Trend)] - [Revised Budget]

-- budget line for trend charts: only in months where the project already reports a forecast
Revised Budget (Trend) = IF ( NOT ISBLANK ( [EAC] ), [Revised Budget] )

-- how much higher the CPI trend is than the PM forecast (positive = PM forecast looks optimistic)
Forecast Gap = [EAC (CPI Trend)] - [EAC]

-- change orders submitted but not approved by the as-of date (exposure / possible recovery)
Pending CO Cost =
VAR d = [As Of Date]
RETURN
    CALCULATE (
        SUM ( fact_change_orders[cost_amount] ),
        REMOVEFILTERS ( dim_date ),
        fact_change_orders[submitted_date] <= d,
        FILTER (
            ALL ( fact_change_orders[status], fact_change_orders[approved_date] ),
            fact_change_orders[status] = "Pending"
                || ( fact_change_orders[status] = "Approved" && fact_change_orders[approved_date] > d )
        )
    ) + 0

Pending CO Price =
VAR d = [As Of Date]
RETURN
    CALCULATE (
        SUM ( fact_change_orders[price_amount] ),
        REMOVEFILTERS ( dim_date ),
        fact_change_orders[submitted_date] <= d,
        FILTER (
            ALL ( fact_change_orders[status], fact_change_orders[approved_date] ),
            fact_change_orders[status] = "Pending"
                || ( fact_change_orders[status] = "Approved" && fact_change_orders[approved_date] > d )
        )
    ) + 0

Projects = COUNTROWS ( VALUES ( dim_project[project_id] ) )

-- more than 1% over budget, so rounding-level noise does not count
Projects Over Budget =
COUNTROWS ( FILTER ( VALUES ( dim_project[project_id] ), [Variance at Completion %] < -0.01 ) ) + 0
```

## 8. Contract and margin (folder: `4 Margin`)

```dax
Original Contract = SUM ( dim_project[original_contract_value] )

Approved CO Revenue =
VAR d = [As Of Date]
RETURN
    CALCULATE (
        SUM ( fact_change_orders[price_amount] ),
        fact_change_orders[status] = "Approved",
        REMOVEFILTERS ( dim_date ),
        fact_change_orders[approved_date] <= d
    ) + 0

Revised Contract = [Original Contract] + [Approved CO Revenue]

Original Margin % = DIVIDE ( [Original Contract] - [Original Budget], [Original Contract] )

Projected Gross Profit = [Revised Contract] - [EAC]

Projected Margin % = DIVIDE ( [Projected Gross Profit], [Revised Contract] )

-- margin fade in percentage points (negative = profit is fading)
Margin Fade (pp) = ( [Projected Margin %] - [Original Margin %] ) * 100

Margin Fade Color = IF ( [Margin Fade (pp)] < 0, "#D64545", "#2E9E6A" )
```

## 9. Billing and WIP (folder: `5 WIP`)

Revenue is recognised on the cost-to-cost method (actual cost ÷ estimate at completion), the usual input
method for construction contracts. Over billing is a liability, under billing an asset.

```dax
% Complete (Cost-to-Cost) =
SUMX ( VALUES ( dim_project[project_id] ), MIN ( 1, DIVIDE ( [Actual Cost], [EAC] ) ) * [Revised Contract] )
    / [Revised Contract]

Earned Revenue =
SUMX ( VALUES ( dim_project[project_id] ), [Revised Contract] * MIN ( 1, DIVIDE ( [Actual Cost], [EAC] ) ) )

Billed to Date =
VAR d = [As Of Date]
RETURN
    IF (
        NOT [Is Future Period],
        CALCULATE ( SUM ( fact_billing[billed_amount] ), REMOVEFILTERS ( dim_date ), fact_billing[billing_period_end] <= d )
    )

Retention Held =
VAR d = [As Of Date]
RETURN
    CALCULATE ( SUM ( fact_billing[retention_held] ), REMOVEFILTERS ( dim_date ), fact_billing[billing_period_end] <= d )

Over (Under) Billing = [Billed to Date] - [Earned Revenue]

Overbillings = SUMX ( VALUES ( dim_project[project_id] ), MAX ( 0, [Over (Under) Billing] ) )

Underbillings = SUMX ( VALUES ( dim_project[project_id] ), MAX ( 0, - [Over (Under) Billing] ) )

Gross Profit to Date = [Earned Revenue] - [Actual Cost]

-- cash in (billed less retention) minus cash out (costs), assuming invoices are paid in the month
Cash Position = [Billed to Date] - [Retention Held] - [Actual Cost]
```

## 10. Status and conditional formatting (folder: `6 Status`)

```dax
Cost Status =
VAR vac = [Variance at Completion %]
VAR cpi = [CPI]
RETURN
    SWITCH (
        TRUE (),
        ISBLANK ( [EAC] ), BLANK (),
        vac < -0.03 || cpi < 0.95, "Over budget",
        vac < 0 || cpi < 0.98, "Watch",
        "On budget"
    )

Status Color =
SWITCH ( [Cost Status], "Over budget", "#D64545", "Watch", "#F2A93B", "On budget", "#2E9E6A" )

VAC Color = IF ( [Variance at Completion] < 0, "#D64545", "#2E9E6A" )

Overrun Color = IF ( [Overrun vs Budget] > 0, "#D64545", "#2E9E6A" )

Billing Color = IF ( [Over (Under) Billing] < 0, "#C77D0A", "#1B2A4A" )

-- heat map: deeper red the bigger the overrun
VAC Heat Color =
VAR v = [Variance at Completion %]
RETURN
    SWITCH (
        TRUE (),
        ISBLANK ( v ), BLANK (),
        v < -0.10, "#B83232",
        v < -0.05, "#E06666",
        v < -0.01, "#F4B6B6",
        v <= 0.01, "#E6EBF2",
        "#A8DCC0"
    )
```

## 11. Labels (folder: `9 Labels`)

```dax
KPI Contract Sub =
"Bid " & FORMAT ( [Original Contract] / 1e6, "$#,0.0" ) & "M + COs " & FORMAT ( [Approved CO Revenue], "$#,0,K" )

KPI Budget Sub =
"Bid " & FORMAT ( [Original Budget] / 1e6, "$#,0.0" ) & "M + COs " & FORMAT ( [Approved CO Cost], "$#,0,K" )

KPI Actual Sub = FORMAT ( [% Complete], "0%" ) & " complete (physical)"

KPI EAC Sub =
VAR v = [Variance at Completion]
RETURN IF ( v < 0, FORMAT ( -v, "$#,0,K" ) & " over budget", FORMAT ( v, "$#,0,K" ) & " under budget" )

KPI Margin Sub =
"Bid margin " & FORMAT ( [Original Margin %], "0.0%" ) & " (" & FORMAT ( [Margin Fade (pp)], "+0.0;-0.0" ) & " pp)"

KPI CPI Sub = "SPI " & FORMAT ( [SPI], "0.00" ) & " · 1.00 = on plan"

KPI Committed Sub = FORMAT ( [Committed % of Budget], "0%" ) & " of revised budget"

KPI Gap Sub =
VAR g = [Forecast Gap]
RETURN IF ( g > 0, FORMAT ( g, "$#,0,K" ) & " above the PM forecast", FORMAT ( -g, "$#,0,K" ) & " below the PM forecast" )

Title Project =
VAR p = SELECTEDVALUE ( dim_project[project_name], "Select one project" )
VAR pm = SELECTEDVALUE ( dim_pm[pm_name] )
VAR st = SELECTEDVALUE ( dim_project[status] )
VAR c = SELECTEDVALUE ( dim_project[client] )
RETURN IF ( HASONEVALUE ( dim_project[project_id] ), p & "  ·  " & c & "  ·  PM " & pm & "  ·  " & st, p )

-- one line per row of the board_points table (1, 2, 3): the summary on the Board Pack slide
Board Point =
VAR worst = TOPN ( 1, VALUES ( dim_project[project_name] ), [Variance at Completion], ASC )
VAR worstName = MAXX ( worst, dim_project[project_name] )
VAR worstVac = CALCULATE ( [Variance at Completion], worst )
RETURN
    SWITCH (
        SELECTEDVALUE ( board_points[Point] ),
        1, [Projects Over Budget] & " of " & [Projects] & " projects are forecast more than 1% over budget. Largest overrun: "
            & worstName & " (" & FORMAT ( -worstVac, "$#,0,K" ) & ").",
        2, "Projected margin is " & FORMAT ( [Projected Margin %], "0.0%" ) & " against " & FORMAT ( [Original Margin %], "0.0%" )
            & " at bid: " & FORMAT ( [Projected Gross Profit] / 1e6, "$#,0.0" ) & "M of gross profit on "
            & FORMAT ( [Revised Contract] / 1e6, "$#,0.0" ) & "M of contracts.",
        3, "The CPI trend is " & FORMAT ( [Forecast Gap], "$#,0,K" ) & " above the PM forecasts, and "
            & FORMAT ( [Pending CO Price], "$#,0,K" ) & " of change orders still wait for client approval."
    )
```

`board_points` is a small calculated table with one column, `Point` = 1, 2, 3:

```dax
board_points = SELECTCOLUMNS ( GENERATESERIES ( 1, 3, 1 ), "Point", [Value] )
```
