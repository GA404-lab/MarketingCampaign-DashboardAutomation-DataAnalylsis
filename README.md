# Marketing Performance Analysis - Excel & Looker Studio
<p align="center">
<img width="1874" height="666" alt="Navigation-ezgif com-optimize" src="https://github.com/user-attachments/assets/8d916427-6640-4bff-b5a2-74a607d5df5d" />
<img width="1366" height="768" alt="lookerdashboard" src="https://github.com/user-attachments/assets/cbe14680-9291-4353-9c06-a736c2387dbc" />
<img width="1874" height="666" alt="exceldashboard" src="https://github.com/user-attachments/assets/089c68c5-7c8b-4a00-9731-7932b6fdb35b" />
<img width="1694" height="666" alt="EXCEL_EJlkZncgPR-ezgif com-optimize" src="https://github.com/user-attachments/assets/6ff247e4-a1e8-451b-9b2a-85e24a044224" />
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

# Data Cleaning
<p align="center">
<img width="1793" height="661" alt="Uncleaned Datapng" src="https://github.com/user-attachments/assets/6b118199-40ce-4d19-8808-0822702b3ec5" />
</p>

## Steps for cleaning

28 columns, 2,241 rows
Data spans 2012 – 2014


- Created a backup before making any changes
- Checked for invalid birth years: anyone born before 1909 or after 2026
```
=COUNTIF(B2:B2241, "<" & 1909) + COUNTIF(B2:B2241, ">" & 2026)
```
Returned 3. Confirmed as data entry errors, removed all 3 rows
- Converted dataset to a formal Excel Table
- Checked for blank cells across the entire dataset:

```
=COUNTBLANK(RawData)
```
Returned 24

- Built a VBA macro to locate blank cells quickly during cleaning:

```vba
Sub Find_Blanks()
' Keyboard Shortcut: Ctrl+Shift+B
    Range("RawData").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
End Sub
```
<p align="center">
<img width="1920" height="716" alt="CleaningMacroforcleaning-ezgif com-optimize" src="https://github.com/user-attachments/assets/de1eb1b5-fca2-4207-891d-dd92f582fb79" />
</p>
all in the Income column, a vital field. All 24 records deleted.
<br><br>

- Converted Income and all spend (Mnt) columns to currency format
- Rearranged AcceptedCmp columns into chronological order (Campaign 1 → 5)
- Audited all categorical columns for invalid entries. Found and removed:
  - 2 records with Marital Status listed as "YOLO"
  - 2 records with Marital Status listed as "Absurd"
  - 3 records with Marital Status listed as "Alone", changed to "Single", an already existing category
  - 3 records with suspicious ages: 114, 115, and 121. The next highest age in the dataset was 74.
- Standardised the Education column from mixed European/American naming conventions to American standard:

```
=IF([@Education]="Basic", "No Secondary Education", IF([@Education]="2n Cycle", "Master", IF([@Education]="Graduation", "Bachelor's", [@Education])))
```

- Confirmed all six campaign columns contain only binary values (0 or 1):
```
=COUNTIF(RawData[[Campaign 1]:[Campaign 6]], ">" & 1) + COUNTIF(RawData[[Campaign 1]:[Campaign 6]], "<" & 0)
```
Returned 0

- Renamed key columns to more intuitive names.
- Applied data validation to Education, Marital Status, Country, and all campaign columns
- Added named ranges for dynamic referencing throughout the workbook

**Final row count after cleaning: 2,209**
<p align="center">
<img width="1679" height="665" alt="EXCEL_L6HBPe98hT" src="https://github.com/user-attachments/assets/83049fd9-bbc9-4ffc-a971-06b3d8702738" />
</p>

---
# Analysis

Created four columns to support analysis:
- **Children**: total number of children in the household
- **Total Spend**: sum of all spend category columns per customer
- **Total Campaigns**: number of campaign offers accepted
- **Age**: calculated as 2014 minus birth year

---

## Which customer segments accept the most campaign offers?

Built pivot tables for each demographic variable (Education, Marital Status, Country, Age) against Total Campaigns accepted. Ran a chi-square goodness of fit test for each variable to identify statistically significant relationships:

```
Expected:  =SUM($C$16:$C$22) * (D16 / SUM($D$16:$D$22))
P-Value:   =CHISQ.TEST(C16:C22, E16:E22)
```

Applied conditional formatting to flag p-values above 0.05. <br>
Mexico was removed from the Goodness of Fit Test for Country due to an extremely low sample size (3).<br>
Education, Marital Status, and Age, all received p-values under 0.05.<br>

For Education and Marital Status I ran a pairwise (1-on-1) goodness of fit tests between individual category pairs. The **Benjamini-Hochberg FDR method** was used to reduce Type I error inflation from multiple comparisons:

```
FDR Threshold:      =(RANK(C28,$C$28:$C$33,1) / COUNT($C$28:$C$33)) * 0.05
Significance Check: =IF(C28 < D28, "Significant", IF(E29 = "Significant", "Significant", "NO"))
```
<p align="center">
<img width="734" height="495" alt="EXCEL_Gs0j9sWH6J" src="https://github.com/user-attachments/assets/e16021f8-c515-490c-8480-491177a30dc8" />
<img width="1848" height="552" alt="EXCEL_mqKd2mpHTt" src="https://github.com/user-attachments/assets/f5a92b58-ed8c-42d4-ad58-81f191a45b3d" />
</p>


---

## What do campaign acceptors spend their money on?

For each spend category, I used formulas to count how many above-average spenders were also high campaign acceptors, then ran a goodness of fit test against the expected distribution.

```
Actual:   =SUMIFS(Total_Accepted_Offers, Spend_Wine, ">" & AVERAGE(Spend_Wine))
Count:    =COUNTIF(Spend_Wine, ">" & AVERAGE(Spend_Wine))
Expected: =SUM($C$7:$C$12) * (D7 / SUM($D$7:$D$12))
P-Value:  =CHISQ.TEST(C7:C12, E7:E12)
```
Returned p-value: 0.0005 <br>
**Campaign acceptors spend their money on categories significantly different from campaign rejectors.**

Created a tabular pivot table for spend category and campaign acceptance. I found the correlation between the average of the various spend categories and the average of accepted offers.
```
=CORREL(H16:H39, I16:I39)
```

Applied conditional formatting to display high, medium, and low correlation tiers across all categories.<br>
**Wine had the highest correlation (r = 0.70), Fruit had the lowest (r = 0.25).**

<p align="center">
<img width="1850" height="588" alt="EXCEL_mWgoEc51BK" src="https://github.com/user-attachments/assets/c5a49044-0d44-4d1c-8ec2-1c67c3adbc71" />
</p>


---

## Is it worth sending more offers after the initial campaign was rejected?

Created a table with a colunm that outputs:

1. **True** if a later campaign was accepted after the initial rejection.
2. **False** if no campaign was accepted after the initial rejection.
3. **Not Applicable** if the first campaign was accepted.

```
=IF(RawData[@[Campaign 1]]=0, IF(OR(RawData[@[Campaign 2]]=1, RawData[@[Campaign 3]]=1, RawData[@[Campaign 4]]=1, RawData[@[Campaign 5]]=1, RawData[@[Campaign 6]]=1), TRUE, FALSE), "Not Applicable")
```

- **22%** of customers who rejected the first campaign went on to accept a later one
- Only **7%** of all customers accepted any campaign at all

**Sending campaigns after an initial rejection is a viable strategy. 1 in 5 initial rejectors eventually converted.**

<p align="center">
<img width="1587" height="349" alt="Initial Rejection" src="https://github.com/user-attachments/assets/bb49fc96-3faa-4684-bd92-20a82312b684" />
</p>

---

## Do acceptors and rejectors prefer different purchase channels?

Built a contingency table comparing above-average channel usage between campaign acceptors and rejectors, then ran a chi-square test of independence:

```
Acceptors above avg:  =COUNTIFS(Total_Accepted_Offers, ">" & 0, of_Web_Purchases, ">" & AVERAGE(of_Web_Purchases))
Rejectors above avg:  =COUNTIFS(Total_Accepted_Offers, 0, of_Web_Purchases, ">" & AVERAGE(of_Web_Purchases))
Expected frequencies: =($E5*C$8) / $E$8
P-Value:              =CHISQ.TEST(C5:D7, C11:D13)
```

Returned p = 0.006. <br>
**The difference in preference between acceptors and rejectors is statistically significant. Catalog was the highest performing channel among acceptors whilst store was the lowest.**

<p align="center">
<img width="1707" height="381" alt="channelacceptance" src="https://github.com/user-attachments/assets/764bb193-1fac-4548-a883-8a129582a693" />
</p>

---

## Do acceptors spend more than rejectors?

Create a list of the total spend of acceptors and rejectors using FILTER:

```
=FILTER(TotalSpend, Total_Accepted_Offers>0)
=FILTER(TotalSpend, Total_Accepted_Offers=0)
```

Ran an F-test to check for equal variances before choosing the right t-test:

```
=FTEST(B3#, C3#)
```

**Returned p-value = 1.91E-23**. Variances are significantly unequal. Applied **Welch's t-test** assuming unequal variances:

```
=T.TEST(B3#, C3#, 2, 3)
```

Returned p-value = 2.40E-59 <br>
| Group | Mean Total Spend |
|---|---|
| Acceptors | $997.36 |
| Rejectors | $460.29 |

**Campaign acceptors spend 117% more than rejectors. This difference is statistically significant (p < 0.001).**

Ran a linear regression analysis using Excel's Data Analysis Toolpak, with Total Spend as the predictor and estimated campaign acceptances as the output. <br>
**Returned p-value = 1.8435E-112**.<br>
Used the resulting coefficients to build a live prediction formula:

```
=G26 + (G27 * I6)
```
Where G26 is the intercept and G27 is the slope from the regression output.<br>
Used **Goal Seek** to find the spend values associated with accepting exactly 1 and 2 campaigns, and saved both as **Scenarios**. Built two macro buttons to switch between them:

```vba
Sub Estimate_1()
    Range("I6").Select
    ActiveSheet.Scenarios("Where estimated number of acceptors is 1").Show
End Sub
Sub Estimate_2()
    Range("I6").Select
    ActiveSheet.Scenarios("Where estimated number of acceptors is 2").Show
End Sub
```

<p align="center">
<img width="1827" height="653" alt="TotalspendingAcceptance-ezgif com-optimize" src="https://github.com/user-attachments/assets/7262692a-423c-4f17-b87e-30f9d3f3147f" />
</p>

---

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

# Dashboard, Automation, and Navigation

## Excel Dashboard

Built an interactive dashboard displaying campaign and customer performance across multiple dimensions.

**KPIs:**
- Total Accepted Offers

**Visuals:**
- Marital Status vs Accepted Offers
- Education Status vs Accepted Offers
- Accepted Offers over time
- World map: Accepted Offers by Country
- Donut chart: ratio of accepted to rejected offers
- Treemap: customer spend by category

**Interactive Controls:**
- Year slicer
- Month slicer
- Country slicer

<p align="center">
<img width="1874" height="666" alt="Exceldashboard-ezgif com-optimize (2)" src="https://github.com/user-attachments/assets/986454f3-99a2-42fb-aa12-60762ec3d170" />
</p>

---

## Navigation and Sheet Protection

Built a custom navigation menu on the left side of each sheet for moving between sheets without using the Excel tab panel.

To keep the navigation buttons functional while preventing users from accidentally selecting or editing them, I unlocked all cells across every worksheet then locked only the navigation panel cells. Each sheet was then protected. Users can freely select and interact with any cell on the sheet, but cannot select the navigation panel. The buttons remain fully clickable.

<p align="center">
<img width="1874" height="666" alt="Navigation-ezgif com-optimize" src="https://github.com/user-attachments/assets/e2e22214-631e-40d6-9432-ae8940d4c135" />
</p>

---

## Looker Studio Dashboard

**Live Dashboard:** [View Interactive Dashboard](https://datastudio.google.com/reporting/b225fb33-b549-43b6-aad8-4e64eead2ede)

**Visuals:**
- Accepted Offers by Education Status
- Accepted Offers by Marital Status
- Accepted Offers: Geo Map
- Accepted Offers over time (month by month)
- KPI: Total Accepted Offers

**Interactive Controls:**
- Year filter
- Month filter
- Country filter

<p align="center">
<img width="1367" height="768" alt="chrome_Xl7dpM38nf-ezgif com-optimize" src="https://github.com/user-attachments/assets/adc1ecc8-da80-4e43-bfa0-38bd4702b09b" />
</p>
