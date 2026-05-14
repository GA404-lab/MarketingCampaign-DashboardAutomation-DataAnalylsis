# Marketing Performance Analysis - Excel & Looker Studio
<p align="center">
<img width="1874" height="666" alt="Navigation-ezgif com-optimize" src="https://github.com/user-attachments/assets/8d916427-6640-4bff-b5a2-74a607d5df5d" />
<img width="1366" height="768" alt="lookerdashboard" src="https://github.com/user-attachments/assets/cbe14680-9291-4353-9c06-a736c2387dbc" />
<img width="1874" height="666" alt="exceldashboard" src="https://github.com/user-attachments/assets/089c68c5-7c8b-4a00-9731-7932b6fdb35b" />
</p>

*Gervon Alcide*

<br>
An analysis of a 2-year marketing campaign dataset covering 2,241 customer records across six campaigns. I cleaned and structured the raw data in Excel, performed statistical analysis using chi-square goodness of fit tests, correlation, and regression, built an interactive Excel dashboard with VBA automation and a custom navigation system, and a companion Looker Studio dashboard.



## Setup

**Tools:** Microsoft Excel, Google Looker Studio <br>
**Dataset:** [Maven Analytics Marketing Campaign Results](https://mavenanalytics.io/data-playground/marketing-campaign-results)<br>
**Live Dashboard (Looker Studio):** [View Interactive Dashboard](https://datastudio.google.com/reporting/b225fb33-b549-43b6-aad8-4e64eead2ede)<br>
**Excel File:** [marketing_data.xlsm](https://github.com/user-attachments/files/27718713/marketing_data.xlsm)

---
## Table of Contents

- [Analysis Overview](#analysis-overview)
- [Insights](#insights)
- [Recommendations](#recommendations)
- [Dashboard and Navigation](#dashboard-and-navigation)

[View full documented process](process.md)

---

# Analysis Overview

I structured my analysis around one main question: what separates customers who accept campaigns and those who don't. I split that one question into 5 and answered them using t-tests, FDR, linear regression, correlation, pivot tables, VBA, Goal Seek, and more. You can find the full documented process with formulas and screenshots [here](process.md).

# Insights

1. Campaign acceptors spend **117% more** than rejectors ($997 vs $460), a statistically significant difference (p < 0.001). Campaign acceptance is a reliable proxy for customer value.
2. **Wine spend has the strongest correlation with campaign acceptance (r = 0.70).** It outperforms every other spend category.
3. **22% of customers who rejected the first campaign went on to accept a later one.** Persistence in campaign outreach is worthwhile.
4. Acceptors and rejectors show **significantly different channel preferences** (p = 0.006). Catalog performs highest among acceptors; store performs lowest. **Channel matters when targeting likely acceptors**.
5. A linear regression model can estimate the spend level at which a customer is likely to accept 1 or 2 campaigns, enabling spend-based segmentation for campaign targeting.

### Most Important Insight

Acceptors spend more than twice as much as non-acceptors. That difference is statistically robust across thousands of records. Businesses running campaigns can use acceptance behaviour as a filter for high-value customer identification.

<p align="center">
<img width="274" height="145" alt="EXCEL_BbAqisfjoK" src="https://github.com/user-attachments/assets/e2513473-c143-458e-bdd9-3e1f280dcaea" />
</p>

---

# Recommendations

1. **Prioritise high wine spenders in campaign targeting.** Wine spend is the strongest predictor of campaign acceptance (r = 0.70). It is a more reliable signal than any other spend category.
2. **Do not stop at the first rejection.** 22% of initial rejectors accepted a later campaign. A multi-touch campaign strategy has a measurable return.
3. **Shift budget toward catalog, away from store.** Acceptors and rejectors have significantly different channel preferences. Catalog outperforms store among acceptors.
4. **Use acceptance history to identify high-value customers.** Acceptors spend 117% more on average. Campaign acceptance can identify high-value customers.
5. **Use the regression model to set spend thresholds for targeting.** Using the regression model you can set a spend threshold that a customer must cross to qualify for outreach. This reduces cost whilst keeping high-acceptance customers.

---

# Dashboard and Navigation

## Excel Dashboard

Built an interactive dashboard displaying campaign and customer performance across multiple dimensions. Users are able to filter by date and country.

<p align="center">
<img width="1874" height="666" alt="Exceldashboard-ezgif com-optimize (2)" src="https://github.com/user-attachments/assets/986454f3-99a2-42fb-aa12-60762ec3d170" />
</p>

---

## Navigation and Sheet Protection

Custom navigation panel built into every sheet. Cells are protected to prevent accidental edits while keeping all buttons fully functional.

<p align="center">
<img width="1874" height="666" alt="Navigation-ezgif com-optimize" src="https://github.com/user-attachments/assets/e2e22214-631e-40d6-9432-ae8940d4c135" />
</p>

---

## Looker Studio Dashboard

Companion web dashboard for sharing results without requiring Excel. Filters by date, and country.

**Live Dashboard:** [View Interactive Dashboard](https://datastudio.google.com/reporting/b225fb33-b549-43b6-aad8-4e64eead2ede)

<p align="center">
<img width="1367" height="768" alt="chrome_Xl7dpM38nf-ezgif com-optimize" src="https://github.com/user-attachments/assets/adc1ecc8-da80-4e43-bfa0-38bd4702b09b" />
</p>
