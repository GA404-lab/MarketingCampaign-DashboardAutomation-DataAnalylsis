# Process
<p align="center">
<img width="1694" height="666" alt="EXCEL_EJlkZncgPR-ezgif com-optimize" src="https://github.com/user-attachments/assets/c77bc59e-6fb7-4249-97f1-4893efd0547a" />
</p>

*Gervon Alcide*

# Table of Contents
- [Methodology](#methodology)
- [Data Cleaning](#data-cleaning)
- [Analysis](#analysis)
  - [Which customer segments accept the most campaign offers?](#which-customer-segments-accept-the-most-campaign-offers)
  - [What do campaign acceptors spend their money on?](#what-do-campaign-acceptors-spend-their-money-on)
  - [Is it worth sending more offers after the initial campaign was rejected?](#is-it-worth-sending-more-offers-after-the-initial-campaign-was-rejected)
  - [Do acceptors and rejectors prefer different purchase channels?](#do-acceptors-and-rejectors-prefer-different-purchase-channels)
  - [Do acceptors spend more than rejectors?](#do-acceptors-spend-more-than-rejectors)
  
[Insights, Recommendations and Visualisation](README.md)

# Methodology

This document covers the full technical process behind the analysis. 
For insights, recommendations, and dashboards, [see the README](README.md).

I start by creating a backup then load the data into Excel. I clean everything I can find, then I reference a personal cleaning checklist to ensure nothing is missed.<br>

For analysis, I structured it around one central question: what separates customers who accept campaigns from those who don't. <br>
I broke that into five sub-questions:

1. Which customer segments accept the most campaign offers?
2. What do campaign acceptors spend their money on?
3. Is it worth sending more offers after the initial campaign was rejected?
4. Do acceptors and rejectors prefer different purchase channels?
5. Do acceptors spend more than rejectors?

# Data Cleaning
<p align="center">
<img width="1793" height="661" alt="Uncleaned Datapng" src="https://github.com/user-attachments/assets/6b118199-40ce-4d19-8808-0822702b3ec5" />
</p>

## Steps for cleaning

28 columns, 2,241 rows
Data spans 2012 – 2014

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

Created a table with a column that outputs:

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

Created a list of the total spend of acceptors and rejectors using FILTER:

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
