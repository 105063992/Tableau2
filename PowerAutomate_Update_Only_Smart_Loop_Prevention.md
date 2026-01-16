# Power Automate Flows - Update Only with Smart Loop Prevention

## Overview

This document provides two Power Automate flows:

1. **Flow 1**: SharePoint Trigger → Pull data from Dataflow (update existing items only)
2. **Flow 2**: Hourly Batch → Update SharePoint items that exist (no create)

### Key Features

| Feature | Description |
|---------|-------------|
| Update Only | No new items created - only existing SharePoint items updated |
| Smart Loop Prevention | SyncSource flag resets after 5 minutes allowing future syncs |
| Basic Mode Conditions | Easy to configure without expressions |
| Clear Step Names | Every step has a descriptive name |

---

## Smart Loop Prevention Strategy

### The Problem with Permanent SyncSource Flag

If SyncSource stays as "Dataflow" forever, the item will NEVER sync again when modified.

### The Solution: Time-Based Reset

We use **two fields** in SharePoint:

| Field | Type | Purpose |
|-------|------|---------|
| SyncSource | Text | Tracks who made the last update |
| LastSyncTime | Date/Time | When the last sync happened |

**Logic:**
- When flow updates SharePoint → Set `SyncSource = "Dataflow"` and `LastSyncTime = Now`
- Trigger condition checks: Skip if `SyncSource = "Dataflow"` AND `LastSyncTime` was within last 5 minutes
- After 5 minutes → Item can sync again!

---

## Prerequisites

### Add These Columns to SharePoint List

| Column Name | Type | Purpose |
|-------------|------|---------|
| SyncSource | Single line of text | Tracks update source |
| LastSyncTime | Date and Time | When last synced |
| Dataverse_ID | Single line of text | Match key to Dataverse |

### Index the Dataverse_ID Column

1. SharePoint List → **Settings** → **List settings**
2. **Indexed columns** → **Create new index**
3. Select **Dataverse_ID**

---

# Flow 1: SharePoint Trigger → Update from Dataflow

## Flow Name: `01_SP_Trigger_Update_From_Dataflow`

### Purpose

When a SharePoint item is modified, pull the latest data from Dataflow table and update the SharePoint item. Includes smart loop prevention that resets after 5 minutes.

---

## Complete Flow Structure

```
┌─────────────────────────────────────────────────────────────────┐
│  STEP 1: Trigger_When_SP_Item_Modified                          │
│  (SharePoint - When an item is created or modified)             │
│                                                                  │
│  Trigger Conditions:                                             │
│  • Skip if SyncSource = 'Dataflow' AND LastSyncTime < 5 min ago │
│  • Skip if Dataverse_ID is empty                                 │
└─────────────────────────────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 2: Get_Dataflow_Record                                     │
│  (Dataverse - List rows)                                         │
│  Filter by Dataverse_ID                                          │
└─────────────────────────────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 3: Check_Record_Found                                      │
│  (Condition - Basic Mode)                                        │
│  Check if Dataflow record exists                                 │
└─────────────────────────────────────────────────────────────────┘
                    │                           │
                   YES                         NO
                    │                           │
                    ▼                           ▼
┌────────────────────────────────┐  ┌────────────────────────────┐
│  STEP 4: Update_SP_Item        │  │  (Do nothing - no record   │
│  • Update fields from Dataflow │  │   found in Dataflow)       │
│  • SyncSource = "Dataflow"     │  │                            │
│  • LastSyncTime = utcNow()     │  └────────────────────────────┘
└────────────────────────────────┘
```

---

## Step-by-Step Instructions

### STEP 1: Trigger_When_SP_Item_Modified

**Connector**: SharePoint  
**Action**: When an item is created or modified

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |

**Rename to**: `Trigger_When_SP_Item_Modified`

#### Trigger Settings (CRITICAL)

Click **...** → **Settings**

**Concurrency Control**:
| Setting | Value |
|---------|-------|
| Limit | `On` |
| Degree of Parallelism | `1` |

**Trigger Conditions** - Add these two:

**Condition 1** - Skip recent Dataflow updates (within 5 minutes):
```
@or(not(equals(triggerOutputs()?['body/SyncSource'], 'Dataflow')), less(ticks(triggerOutputs()?['body/LastSyncTime']), ticks(addMinutes(utcNow(), -5))))
```

**Condition 2** - Only run if Dataverse_ID exists:
```
@not(empty(triggerOutputs()?['body/Dataverse_ID']))
```

**Explanation of Condition 1:**
- Runs if SyncSource is NOT "Dataflow" (normal user edit)
- OR if LastSyncTime is older than 5 minutes ago (enough time passed)
- This prevents immediate loop but allows future syncs

Click **Done**

---

### STEP 2: Get_Dataflow_Record

**Connector**: Microsoft Dataverse  
**Action**: List rows

| Property | Value |
|----------|-------|
| Table name | `Sales_Summarize_For_App` (your Dataflow destination table) |

**Show advanced options:**

| Property | Value |
|----------|-------|
| Filter rows | `cr6b0_dataverse_id eq '` then select **Dataverse_ID** from Dynamic content then add `'` |
| Row count | `1` |

**Full Filter rows value:**
```
cr6b0_dataverse_id eq '@{triggerOutputs()?['body/Dataverse_ID']}'
```

**Select columns** (optional):
```
cr6b0_dataverse_id,cr6b0_region,cr6b0_subregion,cr6b0_year,cr6b0_quarter,cr6b0_month,cr6b0_forecastingtype,cr6b0_pacingtype,cr6b0_internalcqreactivesales,cr6b0_externalcqreactivesales,cr6b0_internalcqbacklogsales,cr6b0_externalcqbacklogsales,cr6b0_internalcqbacklog,cr6b0_externalcqbacklog,cr6b0_internalcqfpconvertiblesales,cr6b0_externalcqfpconvertiblesales
```

**Rename to**: `Get_Dataflow_Record`

---

### STEP 3: Check_Record_Found

**Connector**: Control  
**Action**: Condition

**Rename to**: `Check_Record_Found`

#### Basic Mode Configuration:

| Field | How to Set |
|-------|------------|
| Left side | Click in box → **Expression** tab → Type: `length(outputs('Get_Dataflow_Record')?['body/value'])` → Click **OK** |
| Operator | Select: `is greater than` |
| Right side | Type: `0` |

**Visual:**
```
┌─────────────────────────────────────────────────────────────┐
│  length(outputs('Get_Dataflow_Record')?['body/value'])      │
│                                                             │
│           is greater than                    0              │
└─────────────────────────────────────────────────────────────┘
```

---

### STEP 4: Update_SP_Item (Inside "If yes" branch)

**Connector**: SharePoint  
**Action**: Update item

**Rename to**: `Update_SP_Item`

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |
| Id | Select **ID** from Dynamic content (from trigger) |

**Field Mappings:**

| SharePoint Field | Value | How to Set |
|------------------|-------|------------|
| **SyncSource** | `Dataflow` | Type directly (plain text) |
| **LastSyncTime** | Current time | Expression: `utcNow()` |
| Region | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_region']` |
| Sub Region | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_subregion']` |
| Year | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_year']` |
| Quarter | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_quarter']` |
| Month | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_month']` |
| Forecasting Type | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_forecastingtype']` |
| Pacing Type | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_pacingtype']` |
| Internal-CQReactiveSales | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqreactivesales']` |
| External-CQReactiveSales | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqreactivesales']` |
| Internal-CQBacklogSales | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqbacklogsales']` |
| External-CQBacklogSales | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqbacklogsales']` |
| Internal-CQBacklog | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqbacklog']` |
| External-CQBacklog | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqbacklog']` |
| Internal-CQFPConvertibleSales | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqfpconvertiblesales']` |
| External-CQFPConvertibleSales | Dataflow value | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqfpconvertiblesales']` |

---

### "If no" Branch

Leave empty - do nothing if record not found in Dataflow.

---

## Flow 1 Complete!

Save and test the flow.

---

# Flow 2: Hourly Batch Update (SharePoint Items Only)

## Flow Name: `02_Hourly_Update_SP_From_Dataflow`

### Purpose

Every hour, get all SharePoint items and update them with latest data from Dataflow. **Only updates existing items - never creates new ones.**

---

## Complete Flow Structure

```
┌─────────────────────────────────────────────────────────────────┐
│  STEP 1: Trigger_Hourly                                          │
│  (Schedule - Recurrence)                                         │
└─────────────────────────────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 2: Initialize_ProcessedCount                               │
│  (Variables - Initialize variable)                               │
└─────────────────────────────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 3: Initialize_UpdatedCount                                 │
│  (Variables - Initialize variable)                               │
└─────────────────────────────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 4: Get_All_SP_Items                                        │
│  (SharePoint - Get items)                                        │
│  Get all items from SharePoint that have Dataverse_ID            │
└─────────────────────────────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 5: Loop_Through_SP_Items                                   │
│  (Control - Apply to each)                                       │
│  Parallelism: 20                                                 │
│                                                                  │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │  STEP 5a: Get_Dataflow_Record_For_Item                    │  │
│  │  (Dataverse - List rows)                                  │  │
│  │  Filter by current SP item's Dataverse_ID                 │  │
│  └───────────────────────────────────────────────────────────┘  │
│                              │                                   │
│                              ▼                                   │
│  ┌───────────────────────────────────────────────────────────┐  │
│  │  STEP 5b: Check_Dataflow_Record_Exists                    │  │
│  │  (Condition - Basic Mode)                                 │  │
│  └───────────────────────────────────────────────────────────┘  │
│           │                                    │                 │
│          YES                                  NO                 │
│           │                                    │                 │
│           ▼                                    ▼                 │
│  ┌─────────────────────────┐          ┌─────────────────────┐   │
│  │  STEP 5c: Update_SP_    │          │  (Skip - no         │   │
│  │  Item_With_Dataflow     │          │   Dataflow record)  │   │
│  │  • SyncSource=Dataflow  │          │                     │   │
│  │  • LastSyncTime=Now     │          └─────────────────────┘   │
│  └─────────────────────────┘                                    │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│  STEP 6: Send_Summary_Email                                      │
│  (Outlook - Send email)                                          │
└─────────────────────────────────────────────────────────────────┘
```

---

## Step-by-Step Instructions

### STEP 1: Trigger_Hourly

**Connector**: Schedule  
**Action**: Recurrence

| Property | Value |
|----------|-------|
| Interval | `1` |
| Frequency | `Hour` |
| Time Zone | `(UTC+04:00) Abu Dhabi, Muscat` |

**Rename to**: `Trigger_Hourly`

---

### STEP 2: Initialize_ProcessedCount

**Connector**: Variables  
**Action**: Initialize variable

| Property | Value |
|----------|-------|
| Name | `ProcessedCount` |
| Type | `Integer` |
| Value | `0` |

**Rename to**: `Initialize_ProcessedCount`

---

### STEP 3: Initialize_UpdatedCount

**Connector**: Variables  
**Action**: Initialize variable

| Property | Value |
|----------|-------|
| Name | `UpdatedCount` |
| Type | `Integer` |
| Value | `0` |

**Rename to**: `Initialize_UpdatedCount`

---

### STEP 4: Get_All_SP_Items

**Connector**: SharePoint  
**Action**: Get items

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |
| Filter Query | `Dataverse_ID ne null` |
| Top Count | `5000` (or leave empty for default) |

**Rename to**: `Get_All_SP_Items`

**Note**: This only gets items that HAVE a Dataverse_ID (items to sync).

---

### STEP 5: Loop_Through_SP_Items

**Connector**: Control  
**Action**: Apply to each

**Select an output from previous steps**: 
- Click in the box
- Select **value** from Dynamic content (from Get_All_SP_Items)

**Rename to**: `Loop_Through_SP_Items`

**Settings** (click ... → Settings):

| Setting | Value |
|---------|-------|
| Concurrency Control | `On` |
| Degree of Parallelism | `20` |

---

### Inside Loop_Through_SP_Items:

#### STEP 5a: Get_Dataflow_Record_For_Item

**Connector**: Microsoft Dataverse  
**Action**: List rows

| Property | Value |
|----------|-------|
| Table name | `Sales_Summarize_For_App` |

**Show advanced options:**

| Property | Value |
|----------|-------|
| Filter rows | See below |
| Row count | `1` |

**Filter rows**: Click in box, then click **Expression** tab and enter:
```
concat('cr6b0_dataverse_id eq ''', items('Loop_Through_SP_Items')?['Dataverse_ID'], '''')
```

Or use this format in the field directly:
```
cr6b0_dataverse_id eq '@{items('Loop_Through_SP_Items')?['Dataverse_ID']}'
```

**Rename to**: `Get_Dataflow_Record_For_Item`

---

#### STEP 5b: Check_Dataflow_Record_Exists

**Connector**: Control  
**Action**: Condition

**Rename to**: `Check_Dataflow_Record_Exists`

#### Basic Mode Configuration:

| Field | How to Set |
|-------|------------|
| Left side | Click → **Expression** → `length(outputs('Get_Dataflow_Record_For_Item')?['body/value'])` → **OK** |
| Operator | `is greater than` |
| Right side | `0` |

---

#### STEP 5c: Update_SP_Item_With_Dataflow (Inside "If yes")

**Connector**: SharePoint  
**Action**: Update item

**Rename to**: `Update_SP_Item_With_Dataflow`

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |
| Id | Select **ID** from Dynamic content (from Loop_Through_SP_Items - current item) |

**How to get ID**: 
- Click in Id field
- In Dynamic content, look for **ID** under "Loop_Through_SP_Items"
- Or use Expression: `items('Loop_Through_SP_Items')?['ID']`

**Field Mappings:**

| SharePoint Field | Value |
|------------------|-------|
| **SyncSource** | `Dataflow` (plain text) |
| **LastSyncTime** | Expression: `utcNow()` |
| Region | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_region']` |
| Sub Region | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_subregion']` |
| Year | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_year']` |
| Quarter | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_quarter']` |
| Month | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_month']` |
| Forecasting Type | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_forecastingtype']` |
| Pacing Type | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_pacingtype']` |
| Internal-CQReactiveSales | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqreactivesales']` |
| External-CQReactiveSales | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqreactivesales']` |
| Internal-CQBacklogSales | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqbacklogsales']` |
| External-CQBacklogSales | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqbacklogsales']` |
| Internal-CQBacklog | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqbacklog']` |
| External-CQBacklog | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqbacklog']` |
| Internal-CQFPConvertibleSales | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqfpconvertiblesales']` |
| External-CQFPConvertibleSales | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqfpconvertiblesales']` |

---

#### "If no" Branch

Leave empty - skip items that don't have a matching Dataflow record.

---

### STEP 6: Send_Summary_Email (After the Loop)

**Connector**: Office 365 Outlook  
**Action**: Send an email (V2)

**Rename to**: `Send_Summary_Email`

| Property | Value |
|----------|-------|
| To | `your.email@bakerhughes.com` |
| Subject | `✅ Hourly SP Update Complete - @{utcNow()}` |
| Body | See below |

**Body (HTML):**
```html
<h2>Hourly SharePoint Update Summary</h2>

<table border="1" cellpadding="5">
  <tr style="background-color: #1E4E79; color: white;">
    <th>Metric</th>
    <th>Value</th>
  </tr>
  <tr>
    <td>Flow Run Time</td>
    <td>@{utcNow()}</td>
  </tr>
  <tr>
    <td>SharePoint Items Processed</td>
    <td>@{length(outputs('Get_All_SP_Items')?['body/value'])}</td>
  </tr>
  <tr>
    <td>Status</td>
    <td style="color: green;">✅ Completed</td>
  </tr>
</table>

<p><strong>Flow:</strong> 02_Hourly_Update_SP_From_Dataflow</p>
```

---

## Flow 2 Complete!

---

## How Smart Loop Prevention Works

### Timeline Example:

```
TIME        EVENT                           SYNCSOURCE    LASTSYNCTIME    FLOW RUNS?
─────────────────────────────────────────────────────────────────────────────────────
10:00:00    User edits SP item              (empty)       (empty)         ✅ YES
10:00:05    Flow updates SP item            "Dataflow"    10:00:05        
10:00:06    SP item modified (by flow)      "Dataflow"    10:00:05        ❌ NO (< 5 min)
10:05:10    User edits SP item again        "Dataflow"    10:00:05        ✅ YES (> 5 min)
10:05:15    Flow updates SP item            "Dataflow"    10:05:15        
10:05:16    SP item modified (by flow)      "Dataflow"    10:05:15        ❌ NO (< 5 min)
```

### The Magic Condition Explained:

```
@or(
  not(equals(triggerOutputs()?['body/SyncSource'], 'Dataflow')),
  less(ticks(triggerOutputs()?['body/LastSyncTime']), ticks(addMinutes(utcNow(), -5)))
)
```

**Translation:**
- Run if SyncSource is NOT "Dataflow" (user edit)
- OR Run if LastSyncTime is more than 5 minutes ago (enough time passed)

---

## Step Names Summary

### Flow 1: SP Trigger

| Step | Name |
|------|------|
| Trigger | `Trigger_When_SP_Item_Modified` |
| Get Dataverse | `Get_Dataflow_Record` |
| Condition | `Check_Record_Found` |
| Update SP | `Update_SP_Item` |

### Flow 2: Hourly Batch

| Step | Name |
|------|------|
| Trigger | `Trigger_Hourly` |
| Variable 1 | `Initialize_ProcessedCount` |
| Variable 2 | `Initialize_UpdatedCount` |
| Get SP Items | `Get_All_SP_Items` |
| Loop | `Loop_Through_SP_Items` |
| Get Dataverse | `Get_Dataflow_Record_For_Item` |
| Condition | `Check_Dataflow_Record_Exists` |
| Update SP | `Update_SP_Item_With_Dataflow` |
| Email | `Send_Summary_Email` |

---

## Troubleshooting

### Error: "Infinite loop detected"

**Solution**: Check trigger conditions are added correctly

### Error: "LastSyncTime is null"

**Solution**: The first sync won't have LastSyncTime. The condition handles this - it will run if SyncSource is not "Dataflow"

### Error: "Action not found"

**Solution**: Make sure step names match exactly (no spaces, use underscores)

### Flow runs too slowly

**Solution**: Increase parallelism from 20 to 50 in Loop_Through_SP_Items settings

---

## Key Differences from Previous Flows

| Feature | Previous | New |
|---------|----------|-----|
| Create new items | Yes | **No** |
| Loop prevention | Permanent flag | **Time-based (5 min reset)** |
| Sync scope | All Dataverse records | **Only SP items** |
| Conditions | Expression mode | **Basic mode** |
| Step names | Default | **Descriptive names** |

---

## Document Information

| Property | Value |
|----------|-------|
| Version | 2.0 |
| Created | January 2026 |
| Author | Operations Excellence Team |

