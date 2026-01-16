# Power Automate Flows - Final Version with Real Table Names

## Environment Details

| Property | Value |
|----------|-------|
| Environment | IET Cordant OTR |
| SharePoint Site | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| SharePoint List | `Forecast Main List` |
| Dataverse Table | `FAD_Summary_Table_V6S` |
| Match Key | `cr6b0_id` (Dataverse) ↔ `Dataverse_ID` (SharePoint) |

---

## Dataverse Column Names Reference

| Column Purpose | Dataverse Column Name |
|----------------|----------------------|
| ID (Match Key) | `cr6b0_id` |
| Region | `cr6b0_region` |
| Sub Region | `cr6b0_subregion` |
| Year | `cr6b0_year` |
| Quarter | `cr6b0_quarters` |
| Month | `cr6b0_month` |
| Forecasting Type | `cr6b0_forecastingtype` |
| Pacing Type | `cr6b0_pacingtype` |
| Internal CQ Reactive Sales | `cr6b0_internalcqreactivesales` |
| External CQ Reactive Sales | `cr6b0_externalcqreactivesales` |
| Internal CQ Backlog Sales | `cr6b0_internalcqbacklogsales` |
| External CQ Backlog Sales | `cr6b0_externalcqbacklogsales` |
| Internal CQ Backlog | `cr6b0_internalcqbacklog` |
| External CQ Backlog | `cr6b0_externalcqbacklog` |
| Internal CQ FP Convertible Sales | `cr6b0_internalcqfpconvertiblesales` |
| External CQ FP Convertible Sales | `cr6b0_externalcqfpconvertiblesales` |

---

## Prerequisites

### Add These Columns to SharePoint List

| Column Name | Type | Purpose |
|-------------|------|---------|
| SyncSource | Single line of text | Loop prevention |
| LastSyncTime | Date and Time | Track sync time |
| Dataverse_ID | Single line of text | Match key |

---

# Flow 1: SharePoint Trigger → Update from Dataflow

## Flow Name: `01_SP_Trigger_Update_From_Dataflow`

---

## Flow Structure

```
Trigger_When_SP_Item_Modified
    │
    ▼
Get_Dataflow_Record
    │
    ▼
Check_Record_Found
    │
   YES
    │
    ▼
Update_SP_Item
```

---

## Step 1: Trigger_When_SP_Item_Modified

**Connector**: SharePoint  
**Action**: When an item is created or modified

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |

**Rename to**: `Trigger_When_SP_Item_Modified`

### Trigger Settings

Click **...** → **Settings**

**Concurrency Control**:
| Setting | Value |
|---------|-------|
| Limit | `On` |
| Degree of Parallelism | `1` |

**Trigger Conditions** (Add both):

```
@or(not(equals(triggerOutputs()?['body/SyncSource'], 'Dataflow')), less(ticks(triggerOutputs()?['body/LastSyncTime']), ticks(addMinutes(utcNow(), -5))))
```

```
@not(empty(triggerOutputs()?['body/Dataverse_ID']))
```

Click **Done**

---

## Step 2: Get_Dataflow_Record

**Connector**: Microsoft Dataverse  
**Action**: List rows

**Rename to**: `Get_Dataflow_Record`

| Property | Value |
|----------|-------|
| Table name | `FAD_Summary_Table_V6S` |
| Row count | `1` |

**Filter rows**:
```
cr6b0_id eq '@{triggerOutputs()?['body/Dataverse_ID']}'
```

**Select columns**:
```
cr6b0_id,cr6b0_region,cr6b0_subregion,cr6b0_year,cr6b0_quarters,cr6b0_month,cr6b0_forecastingtype,cr6b0_pacingtype,cr6b0_internalcqreactivesales,cr6b0_externalcqreactivesales,cr6b0_internalcqbacklogsales,cr6b0_externalcqbacklogsales,cr6b0_internalcqbacklog,cr6b0_externalcqbacklog,cr6b0_internalcqfpconvertiblesales,cr6b0_externalcqfpconvertiblesales
```

---

## Step 3: Check_Record_Found

**Connector**: Control  
**Action**: Condition

**Rename to**: `Check_Record_Found`

**Basic Mode:**

| Field | Value |
|-------|-------|
| Left side | Expression: `length(outputs('Get_Dataflow_Record')?['body/value'])` |
| Operator | `is greater than` |
| Right side | `0` |

---

## Step 4: Update_SP_Item (Inside "If yes")

**Connector**: SharePoint  
**Action**: Update item

**Rename to**: `Update_SP_Item`

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |
| Id | `@{triggerOutputs()?['body/ID']}` (Dynamic content: ID from trigger) |

### Field Mappings

| SharePoint Field | Value |
|------------------|-------|
| **SyncSource** | `Dataflow` (type directly) |
| **LastSyncTime** | Expression: `utcNow()` |
| Region | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_region']` |
| Sub Region | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_subregion']` |
| Year | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_year']` |
| Quarter | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_quarters']` |
| Month | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_month']` |
| Forecasting Type | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_forecastingtype']` |
| Pacing Type | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_pacingtype']` |
| Internal-CQReactiveSales | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqreactivesales']` |
| External-CQReactiveSales | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqreactivesales']` |
| Internal-CQBacklogSales | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqbacklogsales']` |
| External-CQBacklogSales | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqbacklogsales']` |
| Internal-CQBacklog | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqbacklog']` |
| External-CQBacklog | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqbacklog']` |
| Internal-CQFPConvertibleSales | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_internalcqfpconvertiblesales']` |
| External-CQFPConvertibleSales | Expression: `first(outputs('Get_Dataflow_Record')?['body/value'])?['cr6b0_externalcqfpconvertiblesales']` |

---

## Flow 1 Complete!

---

# Flow 2: Hourly Batch Update (Large List with Pagination)

## Flow Name: `02_Hourly_Update_Large_SP_List`

---

## Complete Flow Structure

```
Trigger_Hourly
    │
    ▼
Initialize_ProcessedCount (Integer = 0)
    │
    ▼
Initialize_UpdatedCount (Integer = 0)
    │
    ▼
Initialize_SkipCount (Integer = 0)
    │
    ▼
Initialize_HasMoreItems (Boolean = true)
    │
    ▼
Initialize_BatchNumber (Integer = 0)
    │
    ▼
Do_Until_All_Items_Processed (until HasMoreItems = false)
    │
    ├── Increment_BatchNumber
    │
    ├── Get_SP_Items_Batch (5000 items, skip SkipCount)
    │
    ├── Check_Items_Returned
    │       │
    │      YES
    │       │
    │       ├── Process_Batch (Apply to Each, Parallelism: 20)
    │       │       │
    │       │       ├── Check_Has_Dataverse_ID
    │       │       │       │
    │       │       │      YES
    │       │       │       │
    │       │       │       ├── Get_Dataflow_Record_For_Item
    │       │       │       │
    │       │       │       ├── Check_Dataflow_Record_Exists
    │       │       │       │       │
    │       │       │       │      YES
    │       │       │       │       │
    │       │       │       │       └── Update_SP_Item_With_Dataflow
    │       │       │       │
    │       │       │       └── (If NO: skip)
    │       │       │
    │       │       └── (If NO: skip)
    │       │
    │       ├── Update_ProcessedCount
    │       │
    │       ├── Update_SkipCount (+5000)
    │       │
    │       └── Check_If_More_Items
    │               │
    │              YES (batch < 5000)
    │               │
    │               └── Set_HasMoreItems_False
    │
    └── (If NO items: Set_HasMoreItems_False_NoItems)
    │
    ▼
Send_Summary_Email
```

---

## Step-by-Step Instructions

### Step 1: Trigger_Hourly

**Connector**: Schedule  
**Action**: Recurrence

| Property | Value |
|----------|-------|
| Interval | `1` |
| Frequency | `Hour` |
| Time Zone | `(UTC+04:00) Abu Dhabi, Muscat` |

**Rename to**: `Trigger_Hourly`

---

### Step 2: Initialize Variables (5 Variables)

#### Initialize_ProcessedCount

| Property | Value |
|----------|-------|
| Name | `ProcessedCount` |
| Type | `Integer` |
| Value | `0` |

#### Initialize_UpdatedCount

| Property | Value |
|----------|-------|
| Name | `UpdatedCount` |
| Type | `Integer` |
| Value | `0` |

#### Initialize_SkipCount

| Property | Value |
|----------|-------|
| Name | `SkipCount` |
| Type | `Integer` |
| Value | `0` |

#### Initialize_HasMoreItems

| Property | Value |
|----------|-------|
| Name | `HasMoreItems` |
| Type | `Boolean` |
| Value | `true` |

#### Initialize_BatchNumber

| Property | Value |
|----------|-------|
| Name | `BatchNumber` |
| Type | `Integer` |
| Value | `0` |

---

### Step 3: Do_Until_All_Items_Processed

**Connector**: Control  
**Action**: Do until

**Rename to**: `Do_Until_All_Items_Processed`

**Condition (Basic Mode):**

| Field | Value |
|-------|-------|
| Left side | Select `HasMoreItems` from Dynamic content |
| Operator | `is equal to` |
| Right side | Select `false` from Expression: `false` |

**Settings:**

| Setting | Value |
|---------|-------|
| Count | `100` |
| Timeout | `PT4H` |

---

### Inside Do Until Loop:

#### Step 3a: Increment_BatchNumber

**Connector**: Variables  
**Action**: Increment variable

| Property | Value |
|----------|-------|
| Name | `BatchNumber` |
| Value | `1` |

**Rename to**: `Increment_BatchNumber`

---

#### Step 3b: Get_SP_Items_Batch

**Connector**: SharePoint  
**Action**: Get items

**Rename to**: `Get_SP_Items_Batch`

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |
| Filter Query | (leave empty) |
| Top Count | `5000` |

**Settings (click ... → Settings):**

| Setting | Value |
|---------|-------|
| Pagination | `On` |
| Threshold | `100000` |

---

#### Step 3c: Check_Items_Returned

**Connector**: Control  
**Action**: Condition

**Rename to**: `Check_Items_Returned`

**Basic Mode:**

| Field | Value |
|-------|-------|
| Left side | Expression: `length(outputs('Get_SP_Items_Batch')?['body/value'])` |
| Operator | `is greater than` |
| Right side | `0` |

---

#### Step 3d: Process_Batch (Inside "If yes")

**Connector**: Control  
**Action**: Apply to each

**Rename to**: `Process_Batch`

**Select an output**: Click in box, then Dynamic content → select **value** from Get_SP_Items_Batch

Or use expression:
```
outputs('Get_SP_Items_Batch')?['body/value']
```

**Settings:**

| Setting | Value |
|---------|-------|
| Concurrency Control | `On` |
| Degree of Parallelism | `20` |

---

##### Inside Process_Batch:

###### Check_Has_Dataverse_ID

**Connector**: Control  
**Action**: Condition

**Rename to**: `Check_Has_Dataverse_ID`

**Basic Mode:**

| Field | Value |
|-------|-------|
| Left side | Expression: `empty(items('Process_Batch')?['Dataverse_ID'])` |
| Operator | `is equal to` |
| Right side | `false` |

---

###### Inside Check_Has_Dataverse_ID - If YES:

**Get_Dataflow_Record_For_Item**

**Connector**: Microsoft Dataverse  
**Action**: List rows

**Rename to**: `Get_Dataflow_Record_For_Item`

| Property | Value |
|----------|-------|
| Table name | `FAD_Summary_Table_V6S` |
| Row count | `1` |

**Filter rows**:
```
cr6b0_id eq '@{items('Process_Batch')?['Dataverse_ID']}'
```

**Select columns**:
```
cr6b0_id,cr6b0_region,cr6b0_subregion,cr6b0_year,cr6b0_quarters,cr6b0_month,cr6b0_forecastingtype,cr6b0_pacingtype,cr6b0_internalcqreactivesales,cr6b0_externalcqreactivesales,cr6b0_internalcqbacklogsales,cr6b0_externalcqbacklogsales,cr6b0_internalcqbacklog,cr6b0_externalcqbacklog,cr6b0_internalcqfpconvertiblesales,cr6b0_externalcqfpconvertiblesales
```

---

**Check_Dataflow_Record_Exists**

**Connector**: Control  
**Action**: Condition

**Rename to**: `Check_Dataflow_Record_Exists`

**Basic Mode:**

| Field | Value |
|-------|-------|
| Left side | Expression: `length(outputs('Get_Dataflow_Record_For_Item')?['body/value'])` |
| Operator | `is greater than` |
| Right side | `0` |

---

###### Inside Check_Dataflow_Record_Exists - If YES:

**Update_SP_Item_With_Dataflow**

**Connector**: SharePoint  
**Action**: Update item

**Rename to**: `Update_SP_Item_With_Dataflow`

| Property | Value |
|----------|-------|
| Site Address | `https://bakerhughes.sharepoint.com/sites/IETBNServicesFinancialAnalytics` |
| List Name | `Forecast Main List` |
| Id | Expression: `items('Process_Batch')?['ID']` |

**Field Mappings:**

| SharePoint Field | Value |
|------------------|-------|
| **SyncSource** | `Dataflow` (type directly) |
| **LastSyncTime** | Expression: `utcNow()` |
| Region | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_region']` |
| Sub Region | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_subregion']` |
| Year | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_year']` |
| Quarter | Expression: `first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_quarters']` |
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

**Increment_UpdatedCount** (After Update action)

**Connector**: Variables  
**Action**: Increment variable

| Property | Value |
|----------|-------|
| Name | `UpdatedCount` |
| Value | `1` |

**Rename to**: `Increment_UpdatedCount`

---

##### End of Process_Batch

---

#### Step 3e: Update_ProcessedCount (After Process_Batch, still inside "If yes" of Check_Items_Returned)

**Connector**: Variables  
**Action**: Set variable

| Property | Value |
|----------|-------|
| Name | `ProcessedCount` |
| Value | Expression: `add(variables('ProcessedCount'), length(outputs('Get_SP_Items_Batch')?['body/value']))` |

**Rename to**: `Update_ProcessedCount`

---

#### Step 3f: Check_If_More_Items

**Connector**: Control  
**Action**: Condition

**Rename to**: `Check_If_More_Items`

**Basic Mode:**

| Field | Value |
|-------|-------|
| Left side | Expression: `length(outputs('Get_SP_Items_Batch')?['body/value'])` |
| Operator | `is less than` |
| Right side | `5000` |

**If YES** (less than 5000 = last batch):

**Set_HasMoreItems_False**

**Connector**: Variables  
**Action**: Set variable

| Property | Value |
|----------|-------|
| Name | `HasMoreItems` |
| Value | `false` |

**Rename to**: `Set_HasMoreItems_False`

---

#### Step 3g: If NO items in Check_Items_Returned

**Set_HasMoreItems_False_NoItems**

**Connector**: Variables  
**Action**: Set variable

| Property | Value |
|----------|-------|
| Name | `HasMoreItems` |
| Value | `false` |

**Rename to**: `Set_HasMoreItems_False_NoItems`

---

### End of Do Until Loop

---

### Step 4: Send_Summary_Email

**Connector**: Office 365 Outlook  
**Action**: Send an email (V2)

**Rename to**: `Send_Summary_Email`

| Property | Value |
|----------|-------|
| To | `mohamed.samir@bakerhughes.com` |
| Subject | `✅ Hourly SP Sync Complete - @{utcNow()}` |

**Body (HTML):**
```html
<h2 style="color: #1E4E79;">SharePoint Sync Summary</h2>

<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%;">
  <tr style="background-color: #1E4E79; color: white;">
    <th>Metric</th>
    <th>Value</th>
  </tr>
  <tr>
    <td>Flow Run Time</td>
    <td>@{utcNow()}</td>
  </tr>
  <tr style="background-color: #f2f2f2;">
    <td>Total Batches</td>
    <td>@{variables('BatchNumber')}</td>
  </tr>
  <tr>
    <td>Total Items Scanned</td>
    <td>@{variables('ProcessedCount')}</td>
  </tr>
  <tr style="background-color: #f2f2f2;">
    <td>Items Updated</td>
    <td>@{variables('UpdatedCount')}</td>
  </tr>
  <tr>
    <td>Status</td>
    <td style="color: green; font-weight: bold;">✅ Completed</td>
  </tr>
</table>

<p><strong>Flow:</strong> 02_Hourly_Update_Large_SP_List</p>
<p><strong>Source:</strong> FAD_Summary_Table_V6S (Dataverse)</p>
<p><strong>Destination:</strong> Forecast Main List (SharePoint)</p>
```

---

## Flow 2 Complete!

---

## Copy-Paste Expressions Reference

### For Flow 1 (SharePoint Trigger)

**Trigger Condition 1:**
```
@or(not(equals(triggerOutputs()?['body/SyncSource'], 'Dataflow')), less(ticks(triggerOutputs()?['body/LastSyncTime']), ticks(addMinutes(utcNow(), -5))))
```

**Trigger Condition 2:**
```
@not(empty(triggerOutputs()?['body/Dataverse_ID']))
```

**Filter rows (Get_Dataflow_Record):**
```
cr6b0_id eq '@{triggerOutputs()?['body/Dataverse_ID']}'
```

**Select columns:**
```
cr6b0_id,cr6b0_region,cr6b0_subregion,cr6b0_year,cr6b0_quarters,cr6b0_month,cr6b0_forecastingtype,cr6b0_pacingtype,cr6b0_internalcqreactivesales,cr6b0_externalcqreactivesales,cr6b0_internalcqbacklogsales,cr6b0_externalcqbacklogsales,cr6b0_internalcqbacklog,cr6b0_externalcqbacklog,cr6b0_internalcqfpconvertiblesales,cr6b0_externalcqfpconvertiblesales
```

**Check Record Found (Condition):**
```
length(outputs('Get_Dataflow_Record')?['body/value'])
```

**Update SharePoint ID:**
```
triggerOutputs()?['body/ID']
```

---

### For Flow 2 (Hourly Batch)

**Filter rows (Get_Dataflow_Record_For_Item):**
```
cr6b0_id eq '@{items('Process_Batch')?['Dataverse_ID']}'
```

**Check Items Returned:**
```
length(outputs('Get_SP_Items_Batch')?['body/value'])
```

**Check Has Dataverse_ID:**
```
empty(items('Process_Batch')?['Dataverse_ID'])
```

**Check Dataflow Record Exists:**
```
length(outputs('Get_Dataflow_Record_For_Item')?['body/value'])
```

**Update SP Item ID:**
```
items('Process_Batch')?['ID']
```

**Update ProcessedCount:**
```
add(variables('ProcessedCount'), length(outputs('Get_SP_Items_Batch')?['body/value']))
```

---

### Field Mapping Expressions (Flow 2)

**Region:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_region']
```

**Sub Region:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_subregion']
```

**Year:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_year']
```

**Quarter:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_quarters']
```

**Month:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_month']
```

**Forecasting Type:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_forecastingtype']
```

**Pacing Type:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_pacingtype']
```

**Internal-CQReactiveSales:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqreactivesales']
```

**External-CQReactiveSales:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqreactivesales']
```

**Internal-CQBacklogSales:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqbacklogsales']
```

**External-CQBacklogSales:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqbacklogsales']
```

**Internal-CQBacklog:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqbacklog']
```

**External-CQBacklog:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqbacklog']
```

**Internal-CQFPConvertibleSales:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_internalcqfpconvertiblesales']
```

**External-CQFPConvertibleSales:**
```
first(outputs('Get_Dataflow_Record_For_Item')?['body/value'])?['cr6b0_externalcqfpconvertiblesales']
```

---

## Step Names Summary

### Flow 1

| Step | Name |
|------|------|
| Trigger | `Trigger_When_SP_Item_Modified` |
| Get Dataverse | `Get_Dataflow_Record` |
| Condition | `Check_Record_Found` |
| Update | `Update_SP_Item` |

### Flow 2

| Step | Name |
|------|------|
| Trigger | `Trigger_Hourly` |
| Variable 1 | `Initialize_ProcessedCount` |
| Variable 2 | `Initialize_UpdatedCount` |
| Variable 3 | `Initialize_SkipCount` |
| Variable 4 | `Initialize_HasMoreItems` |
| Variable 5 | `Initialize_BatchNumber` |
| Main Loop | `Do_Until_All_Items_Processed` |
| Increment | `Increment_BatchNumber` |
| Get SP | `Get_SP_Items_Batch` |
| Check Items | `Check_Items_Returned` |
| Process Loop | `Process_Batch` |
| Check ID | `Check_Has_Dataverse_ID` |
| Get Dataverse | `Get_Dataflow_Record_For_Item` |
| Check Exists | `Check_Dataflow_Record_Exists` |
| Update SP | `Update_SP_Item_With_Dataflow` |
| Increment Updated | `Increment_UpdatedCount` |
| Update Processed | `Update_ProcessedCount` |
| Check More | `Check_If_More_Items` |
| Set False | `Set_HasMoreItems_False` |
| Email | `Send_Summary_Email` |

---

## Document Information

| Property | Value |
|----------|-------|
| Version | FINAL |
| Created | January 16, 2026 |
| Environment | IET Cordant OTR |
| Dataverse Table | FAD_Summary_Table_V6S |
| SharePoint List | Forecast Main List |

