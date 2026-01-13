# Power BI Migration Design
## Global BN OTR Internal Training Dashboard
### Tableau to Power BI Storytelling Migration

---

| **Attribute** | **Value** |
|---------------|-----------|
| Source System | Tableau Server (BHI) |
| Target System | Power BI Service |
| Total Pages | 5 Sub-pages |
| Primary Audiences | Leadership, Training Managers, HR, Operations |

---

# 1. Dashboard Story Flow

This Power BI solution transforms the Global BN OTR Internal Training Dashboard from a single long-scrolling Tableau page into a structured 5-page narrative that guides users from executive summary through detailed analysis. The dashboard tells the story of organizational training effectiveness, starting with a high-level health check of certifications and skills, drilling into session activity and costs, examining competency gaps, and concluding with actionable details for renewals and remediation.

**The logical flow follows:**

1. **Executive Overview** - Are we meeting training goals? What's our certification health?
2. **Training Sessions & Cost** - How are we investing in training? What's the ROI?
3. **Technical Skills & Competency** - Where are our skill gaps? Which courses are most effective?
4. **Certification Renewals & Actions** - What's at risk? Who needs immediate attention?
5. **Full Details** - Employee-level drill-down for operational management

---

# 2. Page Architecture Summary

| Page # | Page Name | Primary Audience | Key Question |
|--------|-----------|------------------|--------------|
| 1 | Executive Overview | Leadership, Executives | What is our overall training health? |
| 2 | Training Sessions & Cost | Training Managers, Finance | How are we investing in training? |
| 3 | Technical Skills & Competency | HR, Training Leads | Where are our competency gaps? |
| 4 | Renewals & Actions | Operations, Quality | What certifications are at risk? |
| 5 | Full Details | Managers, Analysts | Who specifically needs action? |

---

# 3. Page-by-Page Design

## Page 1: Executive Overview

**Purpose:** Provide leadership with a 30-second health check on organizational training and certification status. Answer: Are we on track with training compliance?

**Primary Audience:** C-Suite, Directors, Regional Leaders

**Key Insights:** Total certifications, active vs expired ratio, regional performance, YoY trends, renewal risk exposure

### Visual Layout

#### Visual 1.1: KPI Card Row (4 cards)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Power BI Card (New Card Visual) |
| Purpose | Display headline KPIs: Total Certifications, Active Rate %, Employees Trained, Avg Certifications per Employee |
| Field Mapping | Values: [Total Certifications], [Active Rate %], [Employees Trained], [Avg Certs Per Employee] |
| Conditional Formatting | Active Rate: Green ≥85%, Amber 70-84%, Red <70% |

#### Visual 1.2: Certification Status Donut Chart

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Donut Chart |
| Purpose | Show Active/Expired/Archived certification distribution |
| Field Mapping | Legend: [Certification Status], Values: [Certification Count] |
| Colors | Active=#70AD47, Expired=#C00000, Archived=#808080 |

#### Visual 1.3: Certifications by Region (Clustered Bar)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Clustered Bar Chart |
| Purpose | Compare certification counts across Business Regions |
| Field Mapping | Y-Axis: [Business Region], X-Axis: [Certification Count], Legend: [Certification Status] |
| Sorting | Descending by Total Certification Count |

#### Visual 1.4: Certification Trend (Line Chart)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Line Chart |
| Purpose | Show certification issuance trend over time |
| Field Mapping | X-Axis: [Certification Issued On] (Year-Quarter), Y-Axis: [Certification Count], Secondary Y: [Active Rate %] |
| Reference Lines | Target Line at 85% Active Rate |

#### Visual 1.5: Function Distribution (Treemap)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Treemap |
| Purpose | Show certification distribution by SOD Function |
| Field Mapping | Group: [SOD Function], Values: [Active Certifications] |

#### Slicers (Top Bar)

- Business Region (Dropdown)
- SOD Function (Dropdown)
- Certification Status (Buttons)
- Date Range (Date Slicer - Certification Issued On)

---

## Page 2: Training Sessions & Cost

**Purpose:** Analyze training session activity, attendee enrollment patterns, and training costs. Answer: How efficiently are we investing in workforce development?

**Primary Audience:** Training Managers, Finance, Operations Leaders

**Key Insights:** Session completion rates, cost per training day, VCO savings, regional investment distribution

### Visual Layout

#### Visual 2.1: Cost KPI Cards (4 cards)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Card (New Card Visual) |
| Metrics | Total Sessions, Total Attendees, Total Training Days, Total VCO Savings ($) |
| Conditional Formatting | VCO Savings: Green if positive, shows trend arrow vs PY |

#### Visual 2.2: Sessions by Status (Stacked Column)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Stacked Column Chart |
| Purpose | Show session enrollment status distribution over time |
| Field Mapping | X-Axis: [Session Year-Month], Y-Axis: [Session Count], Legend: [Session Enrollment Status] |

#### Visual 2.3: Training Cost Analysis (Combo Chart)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Line and Clustered Column Chart |
| Purpose | Compare BR Cost, T&L Cost with VCO Savings trend |
| Field Mapping | X-Axis: [Session Start Year], Columns: [BR Cost], [T&L Cost], Line: [Total VCO Savings] |

#### Visual 2.4: Cost by Region (Matrix)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Matrix |
| Purpose | Show cost breakdown by region and function |
| Field Mapping | Rows: [Location Region], Columns: [SOD Function], Values: [Total Cost], [VCO Savings] |
| Conditional Formatting | Data bars on Total Cost, color scale on VCO Savings |

#### Visual 2.5: Course Popularity (Horizontal Bar)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Bar Chart (Top N) |
| Purpose | Show top 10 most attended courses |
| Field Mapping | Y-Axis: [Course Name], X-Axis: [Attendee Count], Filter: Top 10 by Attendees |

---

## Page 3: Technical Skills & Competency

**Purpose:** Analyze technical skill distribution, competency gaps, and certification coverage by training category. Answer: Where are our biggest skill gaps and what courses address them?

**Primary Audience:** HR, Training Leads, Technical Managers

**Key Insights:** Competency coverage %, skill gaps by region/function, certification density heatmap

### Visual Layout

#### Visual 3.1: Skills Summary KPIs (3 cards)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Card |
| Metrics | Unique Skills Tracked, Employees with Certifications, Avg Skills per Employee |

#### Visual 3.2: Competency Heatmap (Matrix)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Matrix with Conditional Formatting |
| Purpose | Show certification coverage by Training Category × Region |
| Field Mapping | Rows: [Training Category], Columns: [Business Region], Values: [Active Certification %] |
| Conditional Formatting | Background Color Scale: Red (<60%) → Yellow (60-80%) → Green (>80%) |

#### Visual 3.3: Skills by Function (Clustered Bar)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Clustered Bar Chart |
| Purpose | Compare certification counts by SOD Function and Status |
| Field Mapping | Y-Axis: [SOD Function], X-Axis: [Certification Count], Legend: [Certification Status] |

#### Visual 3.4: Top Certifications (Table)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Table |
| Purpose | List most common certifications with holder counts |
| Columns | [Certification Title], [Active Count], [Expired Count], [Active %] |
| Conditional Formatting | Active %: Data bars, Red flag icon if <70% |

#### Visual 3.5: Certification Trend by Category (Area Chart)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Stacked Area Chart |
| Purpose | Show certification issuance trend by Training Category over time |
| Field Mapping | X-Axis: [Certification Issued On] (Year-Quarter), Y-Axis: [Count], Legend: [Training Category] |

---

## Page 4: Certification Renewals & Actions

**Purpose:** Identify certifications at risk of expiration, track renewal aging, and prioritize remediation actions. Answer: What certifications need immediate attention and what's our exposure?

**Primary Audience:** Operations Managers, Quality, Training Coordinators

**Key Insights:** Renewal pipeline (30/60/90 days), expired certification count, aging distribution, action priority list

### Visual Layout

#### Visual 4.1: Renewal Risk KPIs (4 cards)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Card |
| Metrics | Expired Certs (Red), Due in 30 Days (Red), Due in 60 Days (Amber), Due in 90 Days (Yellow) |
| Conditional Formatting | Card backgrounds: Expired=Red, 30Days=Red, 60Days=Amber, 90Days=Yellow |

#### Visual 4.2: Renewal Pipeline Funnel

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Funnel Chart |
| Purpose | Show renewal urgency distribution |
| Field Mapping | Category: [Renewal Status], Values: [Certification Count] |
| Sort Order | Expired → 30 Days → 60 Days → 90 Days → 120 Days → Active |

#### Visual 4.3: Renewals by Region (Stacked Bar)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | 100% Stacked Bar Chart |
| Purpose | Compare renewal risk distribution across regions |
| Field Mapping | Y-Axis: [Business Region], X-Axis: [Certification Count %], Legend: [Renewal Status] |

#### Visual 4.4: Action Priority Table

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Table with Conditional Formatting |
| Purpose | Actionable list of employees needing renewal attention |
| Columns | [Employee Name], [Certification Title], [Renewal Date], [Days Until Expiry], [Region], [Manager] |
| Conditional Formatting | Days Until Expiry: Red <0, Red icon <30, Amber icon <60, Yellow icon <90 |
| Filter | Default: Renewal Status NOT IN ('Active'), sorted by Days Until Expiry ASC |

#### Visual 4.5: Expiry Timeline (Line Chart)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Line Chart |
| Purpose | Show upcoming expirations over next 12 months |
| Field Mapping | X-Axis: [Renewal Date] (Month), Y-Axis: [Expiring Certifications Count] |
| Reference Lines | Vertical line at TODAY, average monthly expiry rate |

---

## Page 5: Full Details

**Purpose:** Provide granular employee-level data for operational management and drill-down analysis. Answer: Who specifically needs what action?

**Primary Audience:** Managers, HR Analysts, Training Coordinators

**Key Insights:** Individual certification records, session attendance, employee training history

### Visual Layout

#### Visual 5.1: Comprehensive Filters Panel

Slicers arranged horizontally at top:
- Business Region (Dropdown)
- Location Country (Dropdown, synced)
- SOD Function (Dropdown)
- Certification Status (Buttons)
- Training Category (Dropdown)
- Job Title (Dropdown)
- Manager (Searchable Dropdown)
- Date Range (Between Slicer)

#### Visual 5.2: Employee Certification Details (Matrix)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Matrix (Expandable) |
| Purpose | Hierarchical view of employees with their certifications |
| Rows | [Business Region] → [Location Country] → [Employee Full Name] → [Certification Title] |
| Values | [Certification Count], [Active Count], [Expired Count], [Renewal Date], [Days Until Expiry] |
| Conditional Formatting | Row background: Red if any expired, Amber if renewal <60 days |

#### Visual 5.3: Session Attendance Details (Table)

| Attribute | Configuration |
|-----------|---------------|
| Visual Type | Table |
| Purpose | Detailed session attendance records |
| Columns | [Employee Name], [Session Name], [Course Name], [Start Date], [End Date], [Enrollment Status], [Completion Date] |
| Features | Exportable, searchable, sortable on all columns |

---

# 4. Required DAX Measures

## 4.1 Core Certification Measures

```dax
// Total Certifications
Total Certifications = COUNTROWS(DF_Certifications)
```

```dax
// Active Certifications
Active Certifications = 
    CALCULATE(
        COUNTROWS(DF_Certifications),
        DF_Certifications[Certification Status] = "Active"
    )
```

```dax
// Active Rate %
Active Rate % = 
    VAR _Total = [Total Certifications]
    VAR _Active = [Active Certifications]
    RETURN
        IF(_Total > 0, DIVIDE(_Active, _Total, 0), BLANK())
```

```dax
// Expired Certifications
Expired Certifications = 
    CALCULATE(
        COUNTROWS(DF_Certifications),
        DF_Certifications[Certification Status] = "Expired"
    )
```

```dax
// Employees Trained (Distinct)
Employees Trained = 
    DISTINCTCOUNT(DF_Certifications[Employee ID])
```

```dax
// Avg Certifications Per Employee
Avg Certs Per Employee = 
    VAR _TotalCerts = [Total Certifications]
    VAR _Employees = [Employees Trained]
    RETURN
        IF(_Employees > 0, DIVIDE(_TotalCerts, _Employees, 0), BLANK())
```

## 4.2 Renewal Status Measures

```dax
// Renewal Status Classification (Calculated Column)
Renewal Status = 
    VAR _DaysToRenewal = DATEDIFF(TODAY(), DF_Certifications[Certification To Renew In], DAY)
    RETURN
        SWITCH(
            TRUE(),
            _DaysToRenewal < 0, "Expired",
            _DaysToRenewal <= 30, "Due in 30 Days",
            _DaysToRenewal <= 60, "Due in 60 Days",
            _DaysToRenewal <= 90, "Due in 90 Days",
            _DaysToRenewal <= 120, "Due in 120 Days",
            "Active"
        )
```

```dax
// Due in 30 Days Count
Due in 30 Days = 
    CALCULATE(
        COUNTROWS(DF_Certifications),
        FILTER(
            DF_Certifications,
            VAR _Days = DATEDIFF(TODAY(), DF_Certifications[Certification To Renew In], DAY)
            RETURN _Days > 0 && _Days <= 30
        )
    )
```

```dax
// Due in 60 Days Count
Due in 60 Days = 
    CALCULATE(
        COUNTROWS(DF_Certifications),
        FILTER(
            DF_Certifications,
            VAR _Days = DATEDIFF(TODAY(), DF_Certifications[Certification To Renew In], DAY)
            RETURN _Days > 30 && _Days <= 60
        )
    )
```

```dax
// Due in 90 Days Count
Due in 90 Days = 
    CALCULATE(
        COUNTROWS(DF_Certifications),
        FILTER(
            DF_Certifications,
            VAR _Days = DATEDIFF(TODAY(), DF_Certifications[Certification To Renew In], DAY)
            RETURN _Days > 60 && _Days <= 90
        )
    )
```

```dax
// Days Until Expiry
Days Until Expiry = 
    DATEDIFF(TODAY(), MAX(DF_Certifications[Certification To Renew In]), DAY)
```

## 4.3 Training Session Measures

```dax
// Total Sessions
Total Sessions = 
    DISTINCTCOUNT(DF_Sessions[Session Name])
```

```dax
// Total Attendees
Total Attendees = 
    COUNTROWS(DF_Sessions)
```

```dax
// Total Training Days
Total Training Days = 
    SUM(DF_Sessions[Total Trg Days])
```

```dax
// Total VCO Savings
Total VCO Savings = 
    SUM(DF_Sessions[Total VCO ($)])
```

```dax
// Total BR Cost
Total BR Cost = 
    SUM(DF_Sessions[BRCost])
```

```dax
// Total T&L Cost
Total T&L Cost = 
    SUM(DF_Sessions[T&L])
```

```dax
// Total Cost
Total Cost = 
    SUM(DF_Sessions[TotalCost])
```

## 4.4 Time Intelligence Measures

```dax
// Certifications YTD
Certifications YTD = 
    CALCULATE(
        [Total Certifications],
        DATESYTD(DateTable[Date])
    )
```

```dax
// Certifications PY (Previous Year)
Certifications PY = 
    CALCULATE(
        [Total Certifications],
        SAMEPERIODLASTYEAR(DateTable[Date])
    )
```

```dax
// YoY Growth %
YoY Growth % = 
    VAR _Current = [Total Certifications]
    VAR _PY = [Certifications PY]
    RETURN
        IF(_PY > 0, DIVIDE(_Current - _PY, _PY, 0), BLANK())
```

```dax
// Rolling 12 Months Certifications
Rolling 12M Certifications = 
    CALCULATE(
        [Total Certifications],
        DATESINPERIOD(
            DateTable[Date],
            MAX(DateTable[Date]),
            -12,
            MONTH
        )
    )
```

## 4.5 Variance & Status Measures

```dax
// Variance vs Target (Absolute)
Variance vs Target = 
    VAR _Actual = [Active Rate %]
    VAR _Target = 0.85  -- 85% target
    RETURN
        _Actual - _Target
```

```dax
// RAG Status Flag
RAG Status = 
    VAR _Rate = [Active Rate %]
    RETURN
        SWITCH(
            TRUE(),
            _Rate >= 0.85, "Green",
            _Rate >= 0.70, "Amber",
            "Red"
        )
```

```dax
// Variance vs PY (Absolute)
Variance vs PY = 
    [Total Certifications] - [Certifications PY]
```

## 4.6 Competency & Skills Measures

```dax
// Unique Skills Tracked
Unique Skills = 
    DISTINCTCOUNT(DF_Certifications[Certification Title])
```

```dax
// Competency Coverage %
Competency Coverage % = 
    VAR _EmployeesWithCerts = 
        CALCULATE(
            DISTINCTCOUNT(DF_Certifications[Employee ID]),
            DF_Certifications[Certification Status] = "Active"
        )
    VAR _TotalEmployees = DISTINCTCOUNT(DF_Certifications[Employee ID])
    RETURN
        IF(_TotalEmployees > 0, DIVIDE(_EmployeesWithCerts, _TotalEmployees, 0), BLANK())
```

---

# 5. UX & Build Guidelines

## 5.1 Layout Principles

Each page should follow a consistent top-to-bottom information hierarchy:

1. **Header Row:** Page title, global filters, last refresh timestamp
2. **KPI Row:** 4-5 headline metrics in card visuals
3. **Primary Charts:** 2-3 main visualizations answering the page question
4. **Supporting Detail:** Tables or secondary charts for drill-down

## 5.2 Color Logic

| Purpose | Hex Code | Usage |
|---------|----------|-------|
| Success/Good | #70AD47 | Active certifications, on-target KPIs, positive trends |
| Warning/Caution | #FFC000 | Due in 60-90 days, near-miss targets |
| Critical/Bad | #C00000 | Expired, overdue, missed targets, <30 days |
| Primary Accent | #1F4E79 | Headers, titles, primary data series |
| Secondary | #2E75B6 | Secondary data series, subtitles |
| Neutral | #808080 | Archived, inactive, reference lines |

## 5.3 Do's and Don'ts

### ✅ DO:

- Use consistent fonts (Segoe UI for body, DIN for numbers)
- Maintain whitespace between visuals for readability
- Use tooltips to provide additional context without cluttering visuals
- Enable cross-filtering between visuals on the same page
- Add data labels only where they add clarity
- Use bookmarks for common filter combinations
- Include a 'Last Refreshed' timestamp on every page
- Test on different screen resolutions (1920x1080 minimum)
- Create mobile-optimized layouts for key pages

### ❌ DON'T:

- Exceed 6-7 visuals per page
- Use 3D charts or unnecessary visual effects
- Overload tables with more than 7-8 columns visible
- Use pie charts for more than 5 categories
- Require scrolling to see key KPIs
- Mix too many color palettes on one page
- Hide critical information in drill-throughs
- Use complex calculated columns when measures suffice

---

# 6. Recommended Data Model

The Power BI solution should use a star schema with the following tables:

## 6.1 Fact Tables

| Table Name | Description & Key Fields |
|------------|--------------------------|
| **DF_Certifications** | LMS Certifications data. Keys: Employee ID, Certification Code, Certification Issued On. Measures: Status, Renewal Date |
| **DF_Sessions** | Training Sessions data. Keys: Employee ID, Session Name, Session Start Date. Measures: Costs, Training Days, VCO |

## 6.2 Dimension Tables

| Table Name | Description & Key Fields |
|------------|--------------------------|
| **DateTable** | Standard date dimension with Year, Quarter, Month, Week columns. Mark as Date Table. |
| **DIM_Employee** | Employee attributes: ID, Name, Job Title, Manager, Region, Country, Function |
| **DIM_Course** | Course/Certification attributes: Code, Title, Category, Duration, Training Type |
| **DIM_Region** | Geographic hierarchy: Business Region, Sub Region, Country, State |

## 6.3 Relationships

```
DateTable[Date] ──── (1:*) ──── DF_Certifications[Certification Issued On]
DateTable[Date] ──── (1:*) ──── DF_Sessions[Session Start Date]
DIM_Employee[Employee ID] ──── (1:*) ──── DF_Certifications[Employee ID]
DIM_Employee[Employee ID] ──── (1:*) ──── DF_Sessions[Employee ID]
DIM_Course[Certification Code] ──── (1:*) ──── DF_Certifications[Certification Code]
DIM_Region[Business Region] ──── (1:*) ──── DF_Certifications[Business Region]
```

---

# 7. Assumptions

The following assumptions have been made based on the Tableau workbook analysis:

1. **Data Sources:** The two primary data sources (Cordant LMS Certifications and Cordant LMS Sessions) will be available as Excel files or SQL connections in Power BI.

2. **Target Values:** An 85% Active Certification Rate is assumed as the organizational target. A Targets table can be created for more granular targets by region/function.

3. **Date Table:** A proper date table exists or will be created spanning the full range of certification and session dates.

4. **Row-Level Security:** The Tableau workbook includes UserSecurityVisibilityCheck calculated field. Equivalent RLS rules should be implemented in Power BI using DAX.

5. **Renewal Logic:** Renewal status buckets (30/60/90/120 days) are calculated from TODAY() relative to the Certification To Renew In field.

6. **Headcount Filter:** Active employees are identified by Headcount Status = 'Active'. Terminated employees should be excluded from current state views.

7. **Cost Fields:** VCO, BR Cost, T&L Cost, and TotalCost fields are available in the Sessions data and are in USD.

8. **Certification Status Values:** Based on the Tableau workbook, valid statuses are: Active, Expired, Archived.

---

# 8. Next Steps

1. Create Power Query connections to source data
2. Build DateTable and dimension tables
3. Implement all DAX measures in a dedicated Measures table
4. Create each page following the visual specifications
5. Apply conditional formatting rules
6. Configure cross-filtering and drill-through actions
7. Implement Row-Level Security
8. User Acceptance Testing with stakeholders
9. Deploy to Power BI Service workspace
10. Schedule data refresh

---

*— End of Document —*
