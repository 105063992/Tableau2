// ============================================================================
// FAD QUERIES - CLEANED AND REFACTORED FOR POWER BI DATAFLOWS
// Part 1: Mapping Tables, Reference Queries, and Dimension Tables
// ============================================================================
// Author: Refactored for Dataflow Migration
// Purpose: Production dashboard queries - Functionally identical to original
// ============================================================================


// ============================================================================
// QUERY: Refresh Schedule
// PURPOSE: Tracks data source refresh dates for monitoring data freshness
// SOURCE: SharePoint Excel file containing refresh schedule tracking
// ============================================================================
let
    // Connect to SharePoint Excel workbook containing refresh schedule data
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Data%20Sources%20Refresh%20Schedule%20%20-Do%20Not%20Delete/SMF%20Data%20Sources%20Refresh%20Dates.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to the Refresh Schedule sheet
    RefreshScheduleSheet = Source{[Item="Refresh Schedule",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(RefreshScheduleSheet, [PromoteAllScalars=true]),
    
    // Set initial column types as text for processing
    InitialTypeConversion = Table.TransformColumnTypes(PromotedHeaders, {
        {"#", type text}, 
        {"BO Orders Data", type text}, 
        {"HFM Variable Cost Data Monthly", type text}, 
        {"WORP_Orders", type text}, 
        {"BN AR_Unbilled_Open BMS master file", type text}, 
        {"Deal Machine", type text}
    }),
    
    // Remove the index/sequence column (not needed for analysis)
    RemoveIndexColumn = Table.RemoveColumns(InitialTypeConversion, {"#"}),
    
    // Add null check column to identify rows with no refresh data
    // Business Rule: Row is valid only if at least one data source has a refresh date
    AddNullCheckColumn = Table.AddColumn(RemoveIndexColumn, "Null Check", each 
        if [BO Orders Data] = null 
            and [HFM Variable Cost Data Monthly] = null 
            and [WORP_Orders] = null 
            and [BI Convertible Order] = null 
            and [BN AR_Unbilled_Open BMS master file] = null 
            and [Deal Machine] = null
        then "All Null" 
        else "Ok"
    ),
    
    // Handle any errors in the null check column
    HandleNullCheckErrors = Table.ReplaceErrorValues(AddNullCheckColumn, {{"Null Check", null}}),
    
    // Filter to keep only valid rows (rows with at least one refresh date)
    FilterValidRows = Table.SelectRows(HandleNullCheckErrors, each ([Null Check] = "Ok")),
    
    // Convert refresh date columns to proper date type
    ConvertToDateTypes = Table.TransformColumnTypes(FilterValidRows, {
        {"BO Orders Data", type date}, 
        {"HFM Variable Cost Data Monthly", type date}, 
        {"WORP_Orders", type date}, 
        {"BN AR_Unbilled_Open BMS master file", type date}, 
        {"Deal Machine", type date}
    }),
    
    // Group and aggregate to get the maximum (latest) refresh date for each data source
    // This provides the most recent refresh date across all tracking entries
    AggregateLatestRefreshDates = Table.Group(ConvertToDateTypes, {"Null Check"}, {
        {"BO Orders Data Refresh Date", each List.Max([BO Orders Data]), type nullable date}, 
        {"HFM Variable Cost Data Monthly Refresh Date", each List.Max([HFM Variable Cost Data Monthly]), type nullable date}, 
        {"WORP_Orders Refresh Date", each List.Max([WORP_Orders]), type nullable date}, 
        {"BN AR_Unbilled_Open BMS master file Refresh Date", each List.Max([BN AR_Unbilled_Open BMS master file]), type nullable date}, 
        {"Deal Machine", each List.Max([Deal Machine]), type nullable date}
    })
in
    AggregateLatestRefreshDates


// ============================================================================
// QUERY: Product Lines Map
// PURPOSE: Master mapping table for WBS Product Lines to Product Hierarchy (PH1-PH5)
// SOURCE: SMF Mapping File Master - Product Lines Map sheet
// BUSINESS USE: Standardizes product categorization across all reports
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Product Lines Map sheet
    ProductLinesMapSheet = Source{[Item="Product Lines Map",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text before promoting headers (handles mixed types)
    PreConvertToText = Table.TransformColumnTypes(ProductLinesMapSheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}, 
        {"Column4", type text}, 
        {"Column5", type text}, 
        {"Column6", type text}, 
        {"Column7", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"WBS Product Line", type text}, 
        {"PH1", type text}, 
        {"PH2", type text}, 
        {"PH3", type text}, 
        {"PH4", type text}, 
        {"PH5", type text}, 
        {"Revenue Stream", type text}
    }),
    
    // Remove duplicate WBS Product Lines (keep first occurrence)
    // Note: Original had two consecutive distinct operations - preserving behavior
    RemoveDuplicates = Table.Distinct(SetColumnTypes, {"WBS Product Line"}),
    RemoveDuplicatesSecondPass = Table.Distinct(RemoveDuplicates, {"WBS Product Line"}),
    
    // Filter out null or empty WBS Product Line values
    FilterValidProductLines = Table.SelectRows(RemoveDuplicatesSecondPass, each 
        [WBS Product Line] <> null and [WBS Product Line] <> ""
    )
in
    FilterValidProductLines


// ============================================================================
// QUERY: Internal External Standard Map
// PURPOSE: Maps source system Internal/External values to standardized values
// SOURCE: SMF Mapping File Master - Internal External Standard Map sheet
// BUSINESS USE: Standardizes customer classification across reports
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Internal External Standard Map sheet
    InternalExternalMapSheet = Source{[Item="Internal External Standard Map",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(InternalExternalMapSheet, [PromoteAllScalars=true]),
    
    // Set column types - some columns are type any due to mixed content
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Internal/External ", type text}, 
        {"Sales Report", type text}, 
        {"Order Report", type text}, 
        {"Backlog Report", type any}, 
        {"Convertible Order", type any}, 
        {"As Sold Report", type any}, 
        {"As Sold Margin", type any}, 
        {"PLM", type any}, 
        {"Unbilled BDR", type any}
    }),
    
    // Filter out null or empty standard values
    FilterValidStandardValues = Table.SelectRows(SetColumnTypes, each 
        [#"Standard Internal/External "] <> null and [#"Standard Internal/External "] <> ""
    )
in
    FilterValidStandardValues


// ============================================================================
// QUERY: Countries_DimTBL
// PURPOSE: Dimension table for geographic hierarchy (Region > Sub Region > Country)
// SOURCE: SMF Mapping File Master - Countries_DimTBL sheet
// BUSINESS USE: Provides standardized geographic dimension for all reports
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Countries dimension table sheet
    CountriesDimSheet = Source{[Item="Countries_DimTBL",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text before promoting headers
    PreConvertToText = Table.TransformColumnTypes(CountriesDimSheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Standard Country", type text}
    }),
    
    // Rename columns to simpler names for easier use
    RenameColumns = Table.RenameColumns(SetColumnTypes, {
        {"Standard Region", "Region"}, 
        {"Standard Sub Region", "Sub Region"}, 
        {"Standard Country", "Country"}
    }),
    
    // Remove completely blank rows
    RemoveBlankRows = Table.SelectRows(RenameColumns, each 
        not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))
    ),
    
    // Add composite key for Region-Country combination
    AddRegionCountryKey = Table.AddColumn(RemoveBlankRows, "RegionCountry", each 
        [Region] & "-" & [Country]
    ),
    
    // Filter out null or empty countries
    FilterValidCountries = Table.SelectRows(AddRegionCountryKey, each 
        [Country] <> null and [Country] <> ""
    )
in
    FilterValidCountries


// ============================================================================
// QUERY: Standard Country List -Sales
// PURPOSE: Country mapping specific to Sales Report data
// SOURCE: SMF Mapping File Master
// BUSINESS USE: Maps source Sales Report regions/countries to standard values
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Sales Country mapping sheet
    SalesCountrySheet = Source{[Item="Standard Country List -Sales",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(SalesCountrySheet, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Standard Country", type text}, 
        {"Region-Sales Report", type text}, 
        {"Country-Sales Report", type text}
    }),
    
    // Remove duplicate region/country combinations from Sales Report
    RemoveDuplicates = Table.Distinct(SetColumnTypes, {"Region-Sales Report", "Country-Sales Report"}),
    
    // Filter out null or empty country values
    FilterValidCountries = Table.SelectRows(RemoveDuplicates, each 
        [#"Country-Sales Report"] <> null and [#"Country-Sales Report"] <> ""
    )
in
    FilterValidCountries


// ============================================================================
// QUERY: Standard Country List-as sold
// PURPOSE: Country mapping specific to As Sold Margin WBS Details data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to As Sold Country mapping sheet
    AsSoldCountrySheet = Source{[Item="Standard Country List-as sold",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(AsSoldCountrySheet, [PromoteAllScalars=true]),
    
    // Remove duplicate region/country combinations
    RemoveDuplicates = Table.Distinct(PromotedHeaders, {"Region-As Sold Margin WBS Details", "Country-As Sold Margin WBS Details"}),
    
    // Filter out null or empty country values
    FilterValidCountries = Table.SelectRows(RemoveDuplicates, each 
        [#"Country-As Sold Margin WBS Details"] <> null and [#"Country-As Sold Margin WBS Details"] <> ""
    )
in
    FilterValidCountries


// ============================================================================
// QUERY: Standard Country List-Backlog
// PURPOSE: Country mapping specific to Backlog Report data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Backlog Country mapping sheet
    BacklogCountrySheet = Source{[Item="Standard Country List-Backlog",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(BacklogCountrySheet, [PromoteAllScalars=true]),
    
    // Remove duplicate region/country combinations
    RemoveDuplicates = Table.Distinct(PromotedHeaders, {"Region-Backlog Report", "Country-Backlog Report"}),
    
    // Filter out null or empty country values
    FilterValidCountries = Table.SelectRows(RemoveDuplicates, each 
        [#"Country-Backlog Report"] <> null and [#"Country-Backlog Report"] <> ""
    ),
    
    // Sort by region for consistent ordering
    SortByRegion = Table.Sort(FilterValidCountries, {{"Region-Backlog Report", Order.Ascending}})
in
    SortByRegion


// ============================================================================
// QUERY: Standard Country List-CO
// PURPOSE: Country mapping specific to Convertible Order data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Convertible Order Country mapping sheet
    COCountrySheet = Source{[Item="Standard Country List-CO",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text
    PreConvertToText = Table.TransformColumnTypes(COCountrySheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}, 
        {"Column4", type text}, 
        {"Column5", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Standard Country", type text}, 
        {"Region-BO CQ Convertible Order", type text}, 
        {"BO CQ Country-Convertible Order", type text}
    }),
    
    // Remove duplicate region/country combinations
    RemoveDuplicates = Table.Distinct(SetColumnTypes, {"Region-BO CQ Convertible Order", "BO CQ Country-Convertible Order"}),
    
    // Filter out null or empty country values
    FilterValidCountries = Table.SelectRows(RemoveDuplicates, each 
        [#"BO CQ Country-Convertible Order"] <> null and [#"BO CQ Country-Convertible Order"] <> ""
    )
in
    FilterValidCountries


// ============================================================================
// QUERY: Standard Country List-Orders
// PURPOSE: Country mapping specific to Order Report data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Orders Country mapping sheet
    OrdersCountrySheet = Source{[Item="Standard Country List-Orders",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text
    PreConvertToText = Table.TransformColumnTypes(OrdersCountrySheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}, 
        {"Column4", type text}, 
        {"Column5", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Standard Country", type text}, 
        {"Region-Order Report", type text}, 
        {"Country-Order Report", type text}
    }),
    
    // Remove duplicate region/country combinations
    RemoveDuplicates = Table.Distinct(SetColumnTypes, {"Region-Order Report", "Country-Order Report"}),
    
    // Filter out null or empty country values
    FilterValidCountries = Table.SelectRows(RemoveDuplicates, each 
        [#"Country-Order Report"] <> null and [#"Country-Order Report"] <> ""
    )
in
    FilterValidCountries


// ============================================================================
// QUERY: Standard Country List-VC
// PURPOSE: Country mapping specific to Variable Cost (HFM Entity) data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Variable Cost Country mapping sheet
    VCCountrySheet = Source{[Item="Standard Country List-VC",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(VCCountrySheet, [PromoteAllScalars=true]),
    
    // Remove duplicate HFM Entity values
    RemoveDuplicates = Table.Distinct(PromotedHeaders, {"HFM Entity-Variable Cost"}),
    
    // Filter out null or empty HFM Entity values
    FilterValidEntities = Table.SelectRows(RemoveDuplicates, each 
        [#"HFM Entity-Variable Cost"] <> null and [#"HFM Entity-Variable Cost"] <> ""
    )
in
    FilterValidEntities


// ============================================================================
// QUERY: Standard Country List-WORP
// PURPOSE: Country mapping specific to WORP (Work Order Reporting) data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to WORP Country mapping sheet
    WORPCountrySheet = Source{[Item="Standard Country List-WORP",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text
    PreConvertToText = Table.TransformColumnTypes(WORPCountrySheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}, 
        {"Column4", type text}, 
        {"Column5", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Filter out null P&L region name values
    FilterValidPLRegion = Table.SelectRows(PromotedHeaders, each 
        [#"P&L region name"] <> null and [#"P&L region name"] <> ""
    ),
    
    // Filter out null Standard Region values
    FilterValidStandardRegion = Table.SelectRows(FilterValidPLRegion, each 
        [Standard Region] <> null and [Standard Region] <> ""
    )
in
    FilterValidStandardRegion


// ============================================================================
// QUERY: Standard Country List-Deal Mach
// PURPOSE: Country mapping specific to Deal Machine data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Deal Machine Country mapping sheet
    DealMachCountrySheet = Source{[Item="Standard Country List-Deal Mach",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text
    PreConvertToText = Table.TransformColumnTypes(DealMachCountrySheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}, 
        {"Column4", type text}, 
        {"Column5", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Standard Country", type text}, 
        {"Region", type text}, 
        {"Primary Country", type text}
    }),
    
    // Filter out null Region values
    FilterValidRegion = Table.SelectRows(SetColumnTypes, each 
        [Region] <> null and [Region] <> ""
    ),
    
    // Remove duplicate Region/Primary Country combinations
    RemoveDuplicates = Table.Distinct(FilterValidRegion, {"Region", "Primary Country"}),
    
    // Filter out null Standard Region values
    FilterValidStandardRegion = Table.SelectRows(RemoveDuplicates, each 
        [Standard Region] <> null and [Standard Region] <> ""
    )
in
    FilterValidStandardRegion


// ============================================================================
// QUERY: Standard Country List-CoCd
// PURPOSE: Maps Company Codes to Standard Region/Sub Region/Country
// SOURCE: SMF Mapping File Master - CoCd_Standard sheet
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Company Code mapping sheet
    CoCdSheet = Source{[Item="CoCd_Standard",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(CoCdSheet, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Country", type text}, 
        {"CoCd", type text}
    }),
    
    // Filter out null Standard Region values
    FilterValidRegions = Table.SelectRows(SetColumnTypes, each 
        [Standard Region] <> null and [Standard Region] <> ""
    )
in
    FilterValidRegions


// ============================================================================
// QUERY: VariableCostCategories_Standard
// PURPOSE: Maps variable cost line items to categories and detailed types
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Variable Cost Categories sheet
    VCCategoriesSheet = Source{[Item="VariableCostCategories_Standard",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(VCCategoriesSheet, [PromoteAllScalars=true]),
    
    // Trim whitespace from Detailed Type column
    TrimDetailedType = Table.TransformColumns(PromotedHeaders, {
        {"Detailed Type", Text.Trim, type text}
    })
in
    TrimDetailedType


// ============================================================================
// QUERY: VariableCostCategories_Std_Nam
// PURPOSE: Variable Cost category mapping specific to NAM (North America) region
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to NAM Variable Cost Categories sheet
    NAMVCCategoriesSheet = Source{[Item="VariableCostCategories_Std_Nam",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text
    PreConvertToText = Table.TransformColumnTypes(NAMVCCategoriesSheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"VC Line Item", type text}, 
        {"VC Category", type text}, 
        {"Detailed Type", type text}
    })
in
    SetColumnTypes


// ============================================================================
// QUERY: DM_StandardMapping
// PURPOSE: Maps Deal Machine product names to standard product hierarchy
// SOURCE: SMF Mapping File Master
// BUSINESS USE: Standardizes Deal Machine product categorization
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to DM Standard Mapping sheet
    DMStandardMappingSheet = Source{[Item="DM_StandardMapping",Kind="Sheet"]}[Data],
    
    // Pre-convert all columns to text
    PreConvertToText = Table.TransformColumnTypes(DMStandardMappingSheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}, 
        {"Column4", type text}, 
        {"Column5", type text}, 
        {"Column6", type text}, 
        {"Column7", type text}, 
        {"Column8", type text}, 
        {"Column9", type text}, 
        {"Column10", type text}, 
        {"Column11", type text}, 
        {"Column12", type text}, 
        {"Column13", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Tier 2 (Legacy)", type text}, 
        {"Primary Tier 3", type text}, 
        {"Tier 4", type text}, 
        {"Prod Lev 1", type text}, 
        {"Prod Lev 2", type text}, 
        {"Prod Lev 3", type text}, 
        {"Prod Lev 4", type text}, 
        {"Product Name", type text}, 
        {"Product Description", type text}, 
        {"Revenue Type", type text}, 
        {"AHM Mapping", type text}, 
        {"Core Mapping", type text}, 
        {"Product Family", type text}
    }),
    
    // Standardize Tier 4 values to PH2 codes
    // Business Rule: Convert legacy tier names to standard PH2 codes
    ReplaceTier4Hardware = Table.ReplaceValue(SetColumnTypes, "Hardware", "BNHW", Replacer.ReplaceText, {"Tier 4"}),
    ReplaceTier4Software = Table.ReplaceValue(ReplaceTier4Hardware, "Software", "BNSW", Replacer.ReplaceText, {"Tier 4"}),
    ReplaceTier4Services = Table.ReplaceValue(ReplaceTier4Software, "BN Services", "BHSV", Replacer.ReplaceText, {"Tier 4"}),
    
    // Rename Tier 4 to PH2 for consistency with product hierarchy naming
    RenameToStandardPH2 = Table.RenameColumns(ReplaceTier4Services, {{"Tier 4", "PH2"}})
in
    RenameToStandardPH2


// ============================================================================
// QUERY: Open T&E_Country_Standard
// PURPOSE: Maps project number legends to standard geography for T&E data
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to T&E Standard mapping sheet
    TEStandardSheet = Source{[Item="Open T&E_Standard",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text
    PreConvertToText = Table.TransformColumnTypes(TEStandardSheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}, 
        {"Column4", type text}, 
        {"Column5", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Standard Country", type text}, 
        {"Project No Legend", type text}, 
        {"Service Region (office)", type text}
    }),
    
    // Remove Service Region column (not needed for mapping)
    RemoveServiceRegion = Table.RemoveColumns(SetColumnTypes, {"Service Region (office)"}),
    
    // Duplicate Project No Legend for extraction
    DuplicateProjectNoLegend = Table.DuplicateColumn(RemoveServiceRegion, "Project No Legend", "Project No Legend - Copy"),
    
    // Extract first 3 characters as project legend prefix
    ExtractProjectLegendPrefix = Table.SplitColumn(
        DuplicateProjectNoLegend, 
        "Project No Legend - Copy", 
        Splitter.SplitTextByPositions({0, 3}, false), 
        {"Project No Legend - Copy.1", "Project No Legend - Copy.2"}
    ),
    
    // Set types for split columns
    SetSplitColumnTypes = Table.TransformColumnTypes(ExtractProjectLegendPrefix, {
        {"Project No Legend - Copy.1", type text}, 
        {"Project No Legend - Copy.2", Int64.Type}
    }),
    
    // Remove the numeric portion (not needed)
    RemoveNumericPortion = Table.RemoveColumns(SetSplitColumnTypes, {"Project No Legend - Copy.2"}),
    
    // Sort by project legend prefix
    SortByPrefix = Table.Sort(RemoveNumericPortion, {{"Project No Legend - Copy.1", Order.Ascending}}),
    
    // Remove duplicate prefixes (keep first occurrence)
    RemoveDuplicatePrefixes = Table.Distinct(SortByPrefix, {"Project No Legend - Copy.1"})
in
    RemoveDuplicatePrefixes


// ============================================================================
// QUERY: MetricsList_DimTBL
// PURPOSE: Dimension table defining available metrics with hierarchy
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Metrics List sheet
    MetricsListSheet = Source{[Item="MetricsList_DimTBL",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(MetricsListSheet, [PromoteAllScalars=true]),
    
    // Add Level Key column (constant value 1)
    AddLevelKey = Table.AddColumn(PromotedHeaders, "Level Key", each 1),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(AddLevelKey, {
        {"Level Key", Int64.Type}, 
        {"Metric ID", Int64.Type}, 
        {"Parent ID", Int64.Type}
    }),
    
    // Remove Level Key column (added but not needed in final output)
    RemoveLevelKey = Table.RemoveColumns(SetColumnTypes, {"Level Key"}),
    
    // Rename columns to standard KPI naming convention
    RenameToKPIConvention = Table.RenameColumns(RemoveLevelKey, {
        {"Metric ID", "KPI_ID"}, 
        {"Metric Name", "KPI"}
    }),
    
    // Remove unnecessary columns
    RemoveUnnecessaryColumns = Table.RemoveColumns(RenameToKPIConvention, {"Comments", "Formula"})
in
    RemoveUnnecessaryColumns


// ============================================================================
// QUERY: PacingForcasting_DimTBL
// PURPOSE: Dimension table for Pacing/Forecasting metrics
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Pacing Forecasting Dimension sheet
    PacingForecastingSheet = Source{[Item="PacingForcasting_DimTBL",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PacingForecastingSheet, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Metric ID", Int64.Type}, 
        {"Metric Name", type text}, 
        {"Parent ID", Int64.Type}
    }),
    
    // Rename to standard KPI naming convention
    RenameToKPIConvention = Table.RenameColumns(SetColumnTypes, {
        {"Metric ID", "KPI_ID"}, 
        {"Metric Name", "KPI"}
    })
in
    RenameToKPIConvention


// ============================================================================
// QUERY: InternalExternal_DimTBL
// PURPOSE: Dimension table for Internal/External customer classification
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Internal/External Dimension sheet
    InternalExternalSheet = Source{[Item="InternalExternal_DimTBL",Kind="Sheet"]}[Data],
    
    // Pre-convert to text
    PreConvertToText = Table.TransformColumnTypes(InternalExternalSheet, {{"Column1", type text}}),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set final column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {{"Internal/External ", type text}})
in
    SetColumnTypes


// ============================================================================
// QUERY: OrderTypeCategory_DimTBL
// PURPOSE: Dimension table for Order Type Categories (Reactive, Backlog, etc.)
// SOURCE: Embedded static data
// ============================================================================
let
    // Create dimension table from embedded static values
    Source = Table.FromRows(
        Json.Document(
            Binary.Decompress(
                Binary.FromText("i45WCkpNTC7JLEtVitWJVnJKTM7OyU8Hs/3ySxSccxKLizPTMlNTlGJjAQ==", BinaryEncoding.Base64), 
                Compression.Deflate
            )
        ), 
        let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [#"Order Type Category" = _t]
    ),
    
    // Set column type
    SetColumnType = Table.TransformColumnTypes(Source, {{"Order Type Category", type text}})
in
    SetColumnType


// ============================================================================
// QUERY: Pacing_Country_DimTBL
// PURPOSE: Dimension table for Pacing report geographic hierarchy
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Pacing Country Dimension sheet
    PacingCountrySheet = Source{[Item="Pacing_Country_DimTBL",Kind="Sheet"]}[Data],
    
    // Pre-convert columns to text
    PreConvertToText = Table.TransformColumnTypes(PacingCountrySheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}
    }),
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PreConvertToText, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Standard Region", type text}, 
        {"Standard Sub Region", type text}, 
        {"Standard Country", type text}
    }),
    
    // Filter out null regions
    FilterValidRegions = Table.SelectRows(SetColumnTypes, each 
        [Standard Region] <> null and [Standard Region] <> ""
    ),
    
    // Rename country column for pacing context
    RenameCountryColumn = Table.RenameColumns(FilterValidRegions, {{"Standard Country", "Base Country"}}),
    
    // Add composite key for Region-Country
    AddRegionCountryKey = Table.AddColumn(RenameCountryColumn, "RegionCountry", each 
        [Standard Region] & "-" & [Base Country]
    )
in
    AddRegionCountryKey


// ============================================================================
// QUERY: Pacing_Sorting_DimTBL
// PURPOSE: Dimension table defining sort order for pacing metrics
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Pacing Sorting Dimension sheet
    PacingSortingSheet = Source{[Item="Pacing_Sorting_DimTBL",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(PacingSortingSheet, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Sort_ID", Int64.Type}, 
        {"Metric Name", type text}
    }),
    
    // Filter out null Sort_ID values
    FilterValidSortIDs = Table.SelectRows(SetColumnTypes, each 
        [Sort_ID] <> null and [Sort_ID] <> ""
    )
in
    FilterValidSortIDs


// ============================================================================
// QUERY: MetricSlicer_DimTBL
// PURPOSE: Dimension table for metric slicer functionality
// SOURCE: SMF Mapping File Master
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Metric Slicer Dimension sheet
    MetricSlicerSheet = Source{[Item="MetricSlicer_DimTBL",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(MetricSlicerSheet, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Metric ID", Int64.Type}, 
        {"Metric Name", type text}
    }),
    
    // Remove unnecessary columns
    RemoveUnnecessaryColumns = Table.RemoveColumns(SetColumnTypes, {"Column4", "Comments"})
in
    RemoveUnnecessaryColumns


// ============================================================================
// QUERY: Backlog_EndUser
// PURPOSE: Extracts End User information from Backlog data for mapping
// SOURCE: SMF_Enduser Reports - BN_Backlog.xlsx
// ============================================================================
let
    // Connect to Backlog End User file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF_Enduser%20Reports/BN_Backlog.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to data sheet
    DataSheet = Source{[Item="Sheet0",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(DataSheet, [PromoteAllScalars=true]),
    
    // Set comprehensive column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Report type", type text}, 
        {"Pole Region Key", type text}, 
        {"Pole Region Text", type text}, 
        {"Project Definition", type text}, 
        {"Project Description", type text}, 
        {"Project Responsible Person Text", type text}, 
        {"WBS Element Number", type text}, 
        {"Sales Order Number", Int64.Type}, 
        {"PH2 Product Line Text", type text}, 
        {"WBS Product Line", type text}, 
        {"Sales Order Type Key", type text}, 
        {"Backlg year", Int64.Type}, 
        {"Calculated Revenue Recognition Year Quarter", type text}, 
        {"Backlog period", Int64.Type}, 
        {"Calculated Backlog Date", type date}, 
        {"Customer Sold To Number", Int64.Type}, 
        {"Customer Sold To Name", type text}, 
        {"Internal/External Text", type text}, 
        {"WBS Element Description", type text}, 
        {"Total Backlog Value USD", type number}, 
        {"Total Backlog Value Local", type number}, 
        {"WBS Status Text", type text}, 
        {"WBS Responsible Person Text", type text}, 
        {"Sales Order Created On Date", type datetime}, 
        {"Sales Order Created By text", type text}, 
        {"WBS Functional Area", type text}, 
        {"WBS Project Type", Int64.Type}, 
        {"WBS RA Key", type text}, 
        {"Customer PO Number", type text}, 
        {"Milestone Number", Int64.Type}, 
        {"Milestone Description", type text}, 
        {"Milestone Usage", type text}, 
        {"Overdue Revenue Flag", type text}, 
        {"PH/WBS Product Line Key", Int64.Type}, 
        {"PH3 Family Text", type text}, 
        {"PH4 Sub Line Text", type text}, 
        {"PH5 Model Text", type text}, 
        {"Project / Flow Indicator", type text}, 
        {"Sales Order Sales Office Text", type text}, 
        {"Sales Order Version Number", Int64.Type}, 
        {"WBS MLST Forecast Fixed Date (Rev Rec)", type datetime}, 
        {"Project Responsible Person Number", Int64.Type}, 
        {"BW Status: SYS0", Int64.Type}, 
        {"Snapshot Date", type date}, 
        {"Backlog Indicator", type any}, 
        {"Calculated Backlog Year Quarter", type text}, 
        {"Backlog row number", Int64.Type}, 
        {"Company Code Key", Int64.Type}, 
        {"Company Code Name", type text}, 
        {"Milestone Incremental POC", Int64.Type}, 
        {"WBS Priority", Int64.Type}, 
        {"Backlog Month", Int64.Type}, 
        {"Sales Order Currency", type text}, 
        {"SO Line Number", Int64.Type}, 
        {"Snapshot Flag", type text}, 
        {"WBS Responsible Person Number", Int64.Type}, 
        {"WBS Project or Service", type text}, 
        {"Customer End User Number", Int64.Type}, 
        {"Customer End User Name", type text}, 
        {"Customer End User Country Text", type text}
    }),
    
    // Select only columns needed for End User mapping
    SelectEndUserColumns = Table.SelectColumns(SetColumnTypes, {"WBS Element Number", "Customer End User Name"}),
    
    // Remove duplicate WBS Elements (keep first End User)
    RemoveDuplicateWBS = Table.Distinct(SelectEndUserColumns, {"WBS Element Number"}),
    
    // Rename columns for mapping table
    RenameForMapping = Table.RenameColumns(RemoveDuplicateWBS, {
        {"WBS Element Number", "WBS"}, 
        {"Customer End User Name", "End User"}
    }),
    
    // Filter out null or empty WBS values
    FilterValidWBS = Table.SelectRows(RenameForMapping, each 
        [WBS] <> null and [WBS] <> ""
    )
in
    FilterValidWBS


// ============================================================================
// QUERY: Sales_EndUser
// PURPOSE: Extracts End User information from Sales data for mapping
// SOURCE: SMF_Enduser Reports - Project_sales_end_user_information.xlsx
// ============================================================================
let
    // Connect to Sales End User file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF_Enduser%20Reports/Project_sales_end_user_information.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to data sheet
    DataSheet = Source{[Item="Sheet0",Kind="Sheet"]}[Data],
    
    // Promote first row to headers
    PromotedHeaders = Table.PromoteHeaders(DataSheet, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"local assigned wbs", type text}, 
        {"local end user account number", Int64.Type}, 
        {"local end user country", type text}, 
        {"local end user description", type text}
    }),
    
    // Select only columns needed for End User mapping
    SelectEndUserColumns = Table.SelectColumns(SetColumnTypes, {"local assigned wbs", "local end user description"}),
    
    // Rename columns for mapping table
    RenameForMapping = Table.RenameColumns(SelectEndUserColumns, {
        {"local assigned wbs", "WBS"}, 
        {"local end user description", "End User"}
    })
in
    RenameForMapping


// ============================================================================
// QUERY: EndUserMapTBL
// PURPOSE: Combined End User mapping from Backlog and Sales sources
// SOURCE: Backlog_EndUser, Sales_EndUser queries
// BUSINESS USE: Provides unified End User lookup by WBS Element
// ============================================================================
let
    // Full outer join Backlog and Sales End User data
    Source = Table.NestedJoin(
        Backlog_EndUser, 
        {"WBS"}, 
        Sales_EndUser, 
        {"WBS"}, 
        "Sales_EndUser", 
        JoinKind.FullOuter
    ),
    
    // Expand Sales End User columns
    ExpandSalesEndUser = Table.ExpandTableColumn(Source, "Sales_EndUser", {"WBS", "End User"}, {"Sales_EndUser.WBS", "Sales_EndUser.End User"}),
    
    // Coalesce WBS from both sources (prefer Backlog, fallback to Sales)
    AddCoalescedWBS = Table.AddColumn(ExpandSalesEndUser, "WBS_EndUser", each 
        if [Sales_EndUser.WBS] = null then [WBS] 
        else if [WBS] = null then [Sales_EndUser.WBS] 
        else [WBS]
    ),
    
    // Sort by WBS
    SortByWBS = Table.Sort(AddCoalescedWBS, {{"WBS", Order.Ascending}}),
    
    // Coalesce End User from both sources (prefer Backlog, fallback to Sales)
    AddCoalescedEndUser = Table.AddColumn(SortByWBS, "EndUser_M", each 
        if [Sales_EndUser.WBS] = null then [End User] 
        else if [WBS] = null then [Sales_EndUser.End User] 
        else [End User]
    ),
    
    // Remove intermediate columns, keep only final mapping columns
    RemoveIntermediateColumns = Table.RemoveColumns(AddCoalescedEndUser, {"WBS", "End User", "Sales_EndUser.WBS", "Sales_EndUser.End User"}),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(RemoveIntermediateColumns, {
        {"WBS_EndUser", type text}, 
        {"EndUser_M", type text}
    }),
    
    // Rename to final column name
    RenameEndUserColumn = Table.RenameColumns(SetColumnTypes, {{"EndUser_M", "End User"}}),
    
    // Sort final result
    SortFinalResult = Table.Sort(RenameEndUserColumn, {{"WBS_EndUser", Order.Ascending}}),
    
    // Filter out null or empty WBS values
    FilterValidWBS = Table.SelectRows(SortFinalResult, each 
        [WBS_EndUser] <> null and [WBS_EndUser] <> ""
    ),
    
    // Remove duplicate WBS entries
    RemoveDuplicateWBS = Table.Distinct(FilterValidWBS, {"WBS_EndUser"})
in
    RemoveDuplicateWBS


// ============================================================================
// QUERY: Last Refresh Date
// PURPOSE: Provides timestamp of last data refresh for reporting
// SOURCE: Generated at runtime using DateTime.LocalNow()
// ============================================================================
let
    // Create base table with placeholder
    Source = Table.FromRows(
        Json.Document(
            Binary.Decompress(
                Binary.FromText("i45WSkksSVSKjQUA", BinaryEncoding.Base64), 
                Compression.Deflate
            )
        ), 
        let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]
    ),
    
    // Set column type
    SetColumnType = Table.TransformColumnTypes(Source, {{"Column1", type text}}),
    
    // Add current datetime as refresh timestamp
    AddRefreshTimestamp = Table.AddColumn(SetColumnType, "Last Refresh Date", each DateTime.LocalNow())
in
    AddRefreshTimestamp


// ============================================================================
// QUERY: LASTREFRESH_For_PowerAutomate
// PURPOSE: Formatted refresh timestamp for Power Automate integration
// SOURCE: Generated at runtime
// ============================================================================
let
    // Create base table
    Source = Table.FromRows(
        Json.Document(
            Binary.Decompress(
                Binary.FromText("i45WUlCKjQUA", BinaryEncoding.Base64), 
                Compression.Deflate
            )
        ), 
        let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]
    ),
    
    // Add refresh datetime
    AddRefreshDateTime = Table.AddColumn(Source, "Last Refresh Date", each DateTime.LocalNow()),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(AddRefreshDateTime, {
        {"Column1", type text}, 
        {"Last Refresh Date", type datetime}
    }),
    
    // Remove placeholder column
    RemovePlaceholder = Table.RemoveColumns(SetColumnTypes, {"Column1"}),
    
    // Add formatted text for Power Automate
    AddFormattedText = Table.AddColumn(RemovePlaceholder, "AutomateLasteRefresh", each 
        "Last Refreshed: " & DateTime.ToText([Last Refresh Date], "MM/dd/yyyy HH:mm:ss")
    )
in
    AddFormattedText


// ============================================================================
// QUERY: ContinousRefresh
// PURPOSE: Placeholder for continuous refresh trigger
// SOURCE: Static text (likely used for refresh scheduling)
// ============================================================================
let
    Source = "dateTimeDateTime.LocalNow()"
in
    Source
