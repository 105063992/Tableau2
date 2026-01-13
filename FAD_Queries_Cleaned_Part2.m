// ============================================================================
// FAD QUERIES - CLEANED AND REFACTORED FOR POWER BI DATAFLOWS
// Part 2: Fact Tables and Major Data Transformations
// ============================================================================
// Author: Refactored for Dataflow Migration
// Purpose: Production dashboard queries - Functionally identical to original
// ============================================================================


// ============================================================================
// QUERY: Variable Cost (HFM Data)
// PURPOSE: Process HFM Variable Cost data with complex header merging
// SOURCE: HFM Variable Cost Data Monthly Excel file
// BUSINESS USE: Cost analysis by region, category, and time period
// DEPENDENCIES: Standard Country List-VC, VariableCostCategories_Standard, HFM Data_NAM_For CM
// ============================================================================
let
    // Connect to HFM Variable Cost source file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20HFM%20Sales/Current%20Week/HFM%20Variable%20Cost%20Data%20Monthly.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to HFM sheet
    HFMSheet = Source{[Item="HFM",Kind="Sheet"]}[Data],
    
    // Filter out header/metadata rows that are not actual data
    // Business Rule: Remove system/dimension headers from the HFM export
    FilterMetadataRows = Table.SelectRows(HFMSheet, each 
        [Column3] <> "DT Total" and 
        [Column3] <> "Reference" and 
        [Column3] <> "Periodic" and 
        [Column3] <> "USD_Reporting - USD_Reporting" and 
        [Column3] <> "PL_OGM36Z - Bently SVC TOTAL" and 
        [Column3] <> "Geographies - Total Geographies" and 
        [Column3] <> "Projects - Total Projects" and 
        [Column3] <> "Rptg - All Reporting Data Types" and 
        [Column3] <> "SuppDtl - Supplemental Detail Dimension" and 
        [Column3] <> "Final" and 
        [Column3] <> "Actual - Actual" and 
        [Column3] <> "Acct_Color - Account Color" and 
        [Column3] <> "Functions - Total Functions"
    ),
    
    // Replace null values with placeholder names for header processing
    ReplaceCOCONulls = Table.ReplaceValue(FilterMetadataRows, null, "COCO Name", Replacer.ReplaceValue, {"Column1"}),
    ReplaceAccountNulls = Table.ReplaceValue(ReplaceCOCONulls, null, "Account", Replacer.ReplaceValue, {"Column2"}),
    
    // Convert columns to text for processing
    ConvertToText = Table.TransformColumns(ReplaceAccountNulls, {
        {"Column1", each Text.From(_), type text}, 
        {"Column2", each Text.From(_), type text}
    }),
    
    // Add index for row reference during header merging
    AddIndex = Table.AddIndexColumn(ConvertToText, "Index", 1, 1, Int64.Type),
    
    // Extract first two rows for header construction
    FirstTwoRows = Table.FirstN(AddIndex, 2),
    
    // Transpose to turn rows into columns for merging
    TransposedRows = Table.Transpose(FirstTwoRows),
    
    ConvertHeaderTypes = Table.TransformColumnTypes(TransposedRows, {
        {"Column1", type text}, 
        {"Column2", type text}
    }),
    
    // Merge the two header rows into single combined headers
    // Business Rule: HFM exports have multi-row headers that need combining
    MergedHeaders = Table.CombineColumns(ConvertHeaderTypes, {"Column1", "Column2"}, Combiner.CombineTextByDelimiter(" "), "Merged"),
    
    // Keep only merged headers column
    HeadersOnly = Table.SelectColumns(MergedHeaders, "Merged"),
    
    // Transpose back to row format
    TransposedHeaders = Table.Transpose(HeadersOnly),
    
    // Remove index from original data
    DataWithoutIndex = Table.RemoveColumns(AddIndex, "Index"),
    
    // Skip the first two rows (original headers) from data
    DataWithoutHeaderRows = Table.Skip(DataWithoutIndex, 2),
    
    // Combine new headers with data
    CombinedTable = Table.Combine({TransposedHeaders, DataWithoutHeaderRows}),
    
    // Promote merged headers
    PromotedHeaders = Table.PromoteHeaders(CombinedTable, [PromoteAllScalars=true]),
    
    // Unpivot month columns to create Year-Month attribute rows
    UnpivotedMonths = Table.UnpivotOtherColumns(PromotedHeaders, {"Account Account", """COCO Name"" ""COCO Name"""}, "Attribute", "Value"),
    
    // Filter out invalid attribute values
    FilterValidAttributes = Table.SelectRows(UnpivotedMonths, each ([Attribute] <> " " and [Attribute] <> "1 2")),
    
    // Trim whitespace from COCO Name
    TrimmedCOCOName = Table.TransformColumns(FilterValidAttributes, {{"""COCO Name"" ""COCO Name""", Text.Trim, type text}}),
    
    // Exclude consolidated/placeholder entities
    FilterExcludedEntities = Table.SelectRows(TrimmedCOCOName, each 
        not Text.Contains([#"""COCO Name"" ""COCO Name"""], "BH_OGMEAS - BH_OGMEAS - Industrial Technology")
    ),
    
    // Join with Standard Country List to get geographic hierarchy
    JoinCountryList = Table.NestedJoin(FilterExcludedEntities, {"""COCO Name"" ""COCO Name"""}, #"Standard Country List-VC", {"HFM Entity-Variable Cost"}, "Standard Country List", JoinKind.LeftOuter),
    
    ExpandCountryList = Table.ExpandTableColumn(JoinCountryList, "Standard Country List", {"Standard Region", "Standard Sub Region", "Standard Country"}, {"Standard Region", "Standard Sub Region", "Standard Country"}),
    
    // Reorder for logical column arrangement
    ReorderColumnsGeography = Table.ReorderColumns(ExpandCountryList, {"Standard Region", "Standard Sub Region", "Standard Country", """COCO Name"" ""COCO Name""", "Account Account", "Attribute", "Value"}),
    
    // Join with Variable Cost Categories for cost classification
    JoinCostCategories = Table.NestedJoin(ReorderColumnsGeography, {"Account Account"}, VariableCostCategories_Standard, {"VC Line Item"}, "VariableCostCategories_Standard", JoinKind.LeftOuter),
    
    ExpandCostCategories = Table.ExpandTableColumn(JoinCostCategories, "VariableCostCategories_Standard", {"VC Category", "Detailed Type"}, {"VC Category", "Detailed Type"}),
    
    ReorderColumnsCategory = Table.ReorderColumns(ExpandCostCategories, {"Standard Region", "Standard Sub Region", "Standard Country", """COCO Name"" ""COCO Name""", "Account Account", "VC Category", "Attribute", "Value"}),
    
    // Split Attribute (Year Month) into separate columns
    SplitYearMonth = Table.SplitColumn(ReorderColumnsCategory, "Attribute", Splitter.SplitTextByDelimiter(" ", QuoteStyle.None), {"Attribute.1", "Attribute.2"}),
    
    SetYearMonthTypes = Table.TransformColumnTypes(SplitYearMonth, {{"Attribute.1", Int64.Type}, {"Attribute.2", type text}}),
    
    // Convert month names to month numbers
    ConvertMonthToNumber = Table.TransformColumns(SetYearMonthTypes, {"Attribute.2", each Date.Month(Date.FromText(_ & " 1")), type number}),
    
    RenameYearMonth = Table.RenameColumns(ConvertMonthToNumber, {{"Attribute.1", "Year"}, {"Attribute.2", "Month"}}),
    
    // Format month as two-digit string (e.g., "01", "02")
    FormatMonth = Table.TransformColumns(RenameYearMonth, {"Month", each Text.PadStart(Text.From(_), 2, "0"), type text}),
    
    ConvertYearToText = Table.TransformColumnTypes(FormatMonth, {{"Year", type text}}),
    
    // Create YearMonth composite key (YYYYMM format)
    AddYearMonthKey = Table.AddColumn(ConvertYearToText, "YearMonth", each [Year] & [Month]),
    
    ReorderFinalColumns = Table.ReorderColumns(AddYearMonthKey, {"Standard Region", "Standard Sub Region", "Standard Country", """COCO Name"" ""COCO Name""", "Account Account", "VC Category", "YearMonth", "Year", "Month", "Value"}),
    
    ConvertValueToNumber = Table.TransformColumnTypes(ReorderFinalColumns, {{"Value", type number}}),
    
    // Create category-specific value columns for analysis
    // Business Rule: Each cost category gets its own column for easier aggregation
    AddCBColumn = Table.AddColumn(ConvertValueToNumber, "C&B", each if [VC Category] = "C&B" then [Value] else 0),
    
    AddDirectMaterialColumn = Table.AddColumn(AddCBColumn, "Direct Material", each if [VC Category] = "Direct Material" then [Value] else 0),
    
    AddDirectLaborColumn = Table.AddColumn(AddDirectMaterialColumn, "Direct Labor, Purchased Services and Others", each if [VC Category] = "Direct Labor, Purchased Services and Others" then [Value] else 0),
    
    AddInternalSupportColumn = Table.AddColumn(AddDirectLaborColumn, "Internal Support", each if [VC Category] = "Internal Support" then [Value] else 0),
    
    AddPastDuesColumn = Table.AddColumn(AddInternalSupportColumn, "Past Dues", each if [VC Category] = "Past Dues" then [Value] else 0),
    
    AddTLColumn = Table.AddColumn(AddPastDuesColumn, "T&L", each if [VC Category] = "T&L" then [Value] else 0),
    
    AddFXColumn = Table.AddColumn(AddTLColumn, "Foreign Exchange (FX)", each if [Detailed Type] = "Foreign Exchange (FX)" then [Value] else 0),
    
    SetFXType = Table.TransformColumnTypes(AddFXColumn, {{"Foreign Exchange (FX)", type number}}),
    
    AddCOGSColumn = Table.AddColumn(SetFXType, "Total Cost of Goods Sold", each if [VC Category] = "Total Cost of Goods Sold" then [Value] else 0),
    
    SetCOGSType = Table.TransformColumnTypes(AddCOGSColumn, {{"Total Cost of Goods Sold", type number}}),
    
    // Create YearMonthDate key in YYYYMMDD format (first day of month)
    AddYearMonthDateKey = Table.AddColumn(SetCOGSType, "YearMonthDate", each [Year] & [Month] & "01"),
    
    // Set all category column types
    SetCategoryTypes = Table.TransformColumnTypes(AddYearMonthDateKey, {
        {"C&B", type number}, 
        {"Direct Material", type number}, 
        {"Direct Labor, Purchased Services and Others", type number}, 
        {"Internal Support", type number}, 
        {"Past Dues", type number}, 
        {"T&L", type number}, 
        {"YearMonthDate", type text}
    }),
    
    // Group by key dimensions and sum values
    GroupedByDimensions = Table.Group(SetCategoryTypes, {"Standard Region", "Standard Sub Region", "Standard Country", "YearMonth", "Year", "YearMonthDate"}, {
        {"C&B", each List.Sum([#"C&B"]), type nullable number}, 
        {"Direct Material", each List.Sum([Direct Material]), type nullable number}, 
        {"Direct Labor, Purchased Services and Others", each List.Sum([#"Direct Labor, Purchased Services and Others"]), type nullable number}, 
        {"Internal Support", each List.Sum([Internal Support]), type nullable number}, 
        {"Past Dues", each List.Sum([Past Dues]), type nullable number}, 
        {"T&L", each List.Sum([#"T&L"]), type nullable number}, 
        {"Foreign Exchange (FX)", each List.Sum([#"Foreign Exchange (FX)"]), type nullable number}, 
        {"Total Cost of Goods Sold", each List.Sum([Total Cost of Goods Sold]), type nullable number}
    }),
    
    // Add Total Variable Cost calculated column
    AddTotalVCColumn = Table.AddColumn(GroupedByDimensions, "Total Variable Cost", each 
        [#"C&B"] + [Direct Material] + [#"Direct Labor, Purchased Services and Others"] + 
        [Internal Support] + [Past Dues] + [#"T&L"] + [#"Foreign Exchange (FX)"]
    ),
    
    SetTotalVCType = Table.TransformColumnTypes(AddTotalVCColumn, {{"Total Variable Cost", type number}}),
    
    // Filter out invalid/null regions
    FilterValidRegions = Table.SelectRows(SetTotalVCType, each ([Standard Region] <> null)),
    
    // Exclude NAM region (handled separately by NAM CM file) and placeholder entities
    FilterExcludeNAM = Table.SelectRows(FilterValidRegions, each 
        ([#"""COCO Name"" ""COCO Name"""] <> "LB_M177 - CP0111_OGMEAS_GE Consolidation-Elimination- PLACEHOLDER-NON LEGAL ENTI" and 
         [#"""COCO Name"" ""COCO Name"""] <> "LB_M615 - OG1134_OGMEAS_Baker Hughes US Energy Holdings LLC") and 
        ([Standard Region] <> "NAM")
    ),
    
    // Append NAM CM data (separate source with more granular NAM breakdown)
    AppendNAMData = Table.Combine({FilterExcludeNAM, #"HFM Data_NAM_For CM"})
in
    AppendNAMData


// ============================================================================
// QUERY: HFM Data_NAM_For CM
// PURPOSE: Process NAM (North America) Variable Cost data from separate CM file
// SOURCE: Cost_P10_Rev1 Excel file with NAM sub-region breakdown
// BUSINESS USE: Provides US regional cost breakdown (Gulf, North, West, Canada)
// DEPENDENCIES: VariableCostCategories_Std_Nam, BO Sales Data_Summarize_For NAM
// ============================================================================
let
    // Connect to NAM CM file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/NAM%20CM%20File/Cost_P10_Rev1.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to Raw Data sheet
    RawDataSheet = Source{[Item="Raw Data",Kind="Sheet"]}[Data],
    
    // Skip header row
    RemoveTopRow = Table.Skip(RawDataSheet, 1),
    
    // Promote headers from data
    PromotedHeaders = Table.PromoteHeaders(RemoveTopRow, [PromoteAllScalars=true]),
    
    // Filter to NAM sub-regions only
    FilterNAMSubRegions = Table.SelectRows(PromotedHeaders, each 
        ([#"Sub-regions"] = "Canada" or 
         [#"Sub-regions"] = "Gulf" or 
         [#"Sub-regions"] = "North" or 
         [#"Sub-regions"] = "West")
    ),
    
    // Add Standard Country (United States for all NAM data)
    AddStandardCountry = Table.AddColumn(FilterNAMSubRegions, "Standard Country", each "United States"),
    SetCountryType = Table.TransformColumnTypes(AddStandardCountry, {{"Standard Country", type text}}),
    
    // Add Standard Region
    AddStandardRegion = Table.AddColumn(SetCountryType, "Standard Region", each "NAM"),
    SetRegionType = Table.TransformColumnTypes(AddStandardRegion, {{"Standard Region", type text}}),
    
    // Standardize sub-region names to US regional format
    // Business Rule: Map source sub-region names to standard "US -Region" format
    ReplaceGulf = Table.ReplaceValue(SetRegionType, "Gulf", "US -Gulf", Replacer.ReplaceText, {"Sub-regions"}),
    ReplaceWest = Table.ReplaceValue(ReplaceGulf, "West", "US -West", Replacer.ReplaceText, {"Sub-regions"}),
    ReplaceNorth = Table.ReplaceValue(ReplaceWest, "North", "US -North", Replacer.ReplaceText, {"Sub-regions"}),
    
    // Rename to standard column name
    RenameSubRegion = Table.RenameColumns(ReplaceNorth, {{"Sub-regions", "Standard Sub Region"}}),
    
    // Remove unnecessary source columns
    RemoveSourceColumns = Table.RemoveColumns(RenameSubRegion, {
        "es account l03", "local wbs description", "es account l03 description", 
        "local assigned wbs", "local project number", "local project description", 
        "local cost center", "local cost center description", 
        "local legal entity/company code", "local account", "source system", 
        "local period code", "ds pc busline", "ds prod tier", "local functional area", 
        "es function", "es function description", "parent ship to region code", 
        "parent sales office region description", "parent ship to country description", 
        "local document number", "local sales order number", "local sales order type", 
        "local text header", "es product line", "es product line l04", 
        "es product line l04 description", "es account l06 description", 
        "Revenue, COGS, OVC", "Sales or Variable Costs", "es account", "Cost Category"
    }),
    
    // Join with NAM-specific Variable Cost Categories mapping
    JoinCostCategories = Table.NestedJoin(RemoveSourceColumns, {"es account descripton"}, VariableCostCategories_Std_Nam, {"VC Line Item"}, "VariableCostCategories_Std_Nam", JoinKind.LeftOuter),
    
    ExpandCostCategories = Table.ExpandTableColumn(JoinCostCategories, "VariableCostCategories_Std_Nam", {"VC Line Item", "VC Category", "Detailed Type"}, {"VC Line Item", "VC Category", "Detailed Type"}),
    
    // Remove additional unused columns
    RemoveAdditionalColumns = Table.RemoveColumns(ExpandCostCategories, {
        "Total Sales Reporting Color", "Total Sales Reporting Color (MM)", 
        "local period quarter", "Monthly local posted date", "local work order", 
        "local user name", "es account descripton"
    }),
    
    // Reorder and rename value column
    ReorderColumns = Table.ReorderColumns(RemoveAdditionalColumns, {"VC Category", "Total usd activity", "Monthly local created date", "es company", "Standard Country", "Standard Region", "VC Line Item", "Detailed Type"}),
    RenameValueColumn = Table.RenameColumns(ReorderColumns, {{"Total usd activity", "Value"}}),
    
    // Convert date column for year/month extraction
    SetDateType = Table.TransformColumnTypes(RenameValueColumn, {{"Monthly local created date", type date}}),
    
    // Extract Year from date
    DuplicateDateForYear = Table.DuplicateColumn(SetDateType, "Monthly local created date", "Monthly local created date - Copy"),
    ExtractYear = Table.TransformColumns(DuplicateDateForYear, {{"Monthly local created date - Copy", Date.Year, Int64.Type}}),
    RenameYear = Table.RenameColumns(ExtractYear, {{"Monthly local created date - Copy", "Year"}}),
    
    // Extract Month from date
    DuplicateDateForMonth = Table.DuplicateColumn(RenameYear, "Monthly local created date", "Monthly local created date - Copy"),
    ExtractMonth = Table.TransformColumns(DuplicateDateForMonth, {{"Monthly local created date - Copy", Date.Month, Int64.Type}}),
    RenameMonth = Table.RenameColumns(ExtractMonth, {{"Monthly local created date - Copy", "Month"}}),
    
    // Create category-specific columns
    AddCBColumn = Table.AddColumn(RenameMonth, "C&B", each if [VC Category] = "C&B" then [Value] else null),
    AddDirectMaterialColumn = Table.AddColumn(AddCBColumn, "Direct Material", each if [VC Category] = "Direct Material" then [Value] else null),
    AddDirectLaborColumn = Table.AddColumn(AddDirectMaterialColumn, "Direct Labor, Purchased Services and Others", each if [VC Category] = "Direct Labor, Purchased Services and Others" then [Value] else null),
    AddInternalSupportColumn = Table.AddColumn(AddDirectLaborColumn, "Internal Support", each if [VC Category] = "Internal Support" then [Value] else null),
    AddTLColumn = Table.AddColumn(AddInternalSupportColumn, "T&L", each if [VC Category] = "T&L" then [Value] else null),
    AddFXColumn = Table.AddColumn(AddTLColumn, "Foreign Exchange (FX)", each if Text.Contains([Detailed Type], "Foreign Exchange (FX)") then [Value] else null),
    
    // Set column types
    SetCategoryTypes = Table.TransformColumnTypes(AddFXColumn, {
        {"Year", type text}, 
        {"Month", type text}, 
        {"C&B", type number}, 
        {"Direct Material", type number}, 
        {"Direct Labor, Purchased Services and Others", type number}, 
        {"Internal Support", type number}, 
        {"T&L", type number}, 
        {"Foreign Exchange (FX)", type number}
    }),
    
    // Create YearMonthDate key
    AddYearMonthDateKey = Table.AddColumn(SetCategoryTypes, "YearMonthDate", each [Year] & [Month] & "01"),
    
    // Append Sales summaries for NAM
    AppendSalesSummaries = Table.Combine({AddYearMonthDateKey, #"BO Sales Data_Summarize_For NAM", #"BO Sales Data_Summarize_For NAM_total Revenue"}),
    
    // Final type conversions
    SetValueType = Table.TransformColumnTypes(AppendSalesSummaries, {{"Value", type number}}),
    RemoveCompanyCode = Table.RemoveColumns(SetValueType, {"es company"}),
    SetFinalTypes = Table.TransformColumnTypes(RemoveCompanyCode, {
        {"Standard Region", type text}, 
        {"YearMonthDate", type text}, 
        {"Value", type number}
    })
in
    SetFinalTypes


// ============================================================================
// QUERY: VariableCostCategories_Std_Nam
// PURPOSE: NAM-specific Variable Cost category mapping
// SOURCE: SMF Mapping File Master - VariableCostCategories_Std_Nam sheet
// BUSINESS USE: Maps NAM account descriptions to standard VC categories
// ============================================================================
let
    // Connect to master mapping file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to NAM VC Categories sheet
    VCCategoriesSheet = Source{[Item="VariableCostCategories_Std_Nam",Kind="Sheet"]}[Data],
    
    // Convert to text before promoting headers
    ConvertToText = Table.TransformColumnTypes(VCCategoriesSheet, {
        {"Column1", type text}, 
        {"Column2", type text}, 
        {"Column3", type text}
    }),
    
    // Promote headers
    PromotedHeaders = Table.PromoteHeaders(ConvertToText, [PromoteAllScalars=true]),
    
    // Set final types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"VC Line Item", type text}, 
        {"VC Category", type text}, 
        {"Detailed Type", type text}
    })
in
    SetColumnTypes


// ============================================================================
// QUERY: BO Sales 2024
// PURPOSE: Load 2024 Sales data from historical archive
// SOURCE: 2024 Sales Excel file
// BUSINESS USE: Historical sales data for trend analysis
// ============================================================================
let
    // Connect to 2024 Sales file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20TS%20Sales%20-%20Histroical/Current%20Week/Previous%20Year-Dont%20Change/2024%20Sales.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to data sheet
    SalesSheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    
    // Promote headers
    PromotedHeaders = Table.PromoteHeaders(SalesSheet, [PromoteAllScalars=true]),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Project Number", type text}, 
        {"Project Name", type text}, 
        {"Region", type text}, 
        {"Country", type text}, 
        {"Internal/External", type text}, 
        {"Sold To", type text}, 
        {"Sold to Description", type text}, 
        {"PM/SM Name", type text}, 
        {"Company Code", Int64.Type}, 
        {"Order Type", type text}, 
        {"Business", type text}, 
        {"Product Line", type text}, 
        {"PHx DS 17 PLs", type text}, 
        {"WBS NUMBER", type text}, 
        {"WBS DESCRIPTION", type text}, 
        {"WBS PRODUCT LINE", type text}, 
        {"MONTH", type text}, 
        {"QTR", Int64.Type}, 
        {"YEAR", Int64.Type}, 
        {"AMOUNT USD", type number}
    }),
    
    // Filter out blank project numbers
    FilterValidProjects = Table.SelectRows(SetColumnTypes, each [Project Number] <> null and [Project Number] <> ""),
    
    // Standardize Internal/External values to match master mapping
    // Business Rule: Map source system values to standard classification
    ReplaceExternalEXT = Table.ReplaceValue(FilterValidProjects, "External (EXT)", "External Customers", Replacer.ReplaceText, {"Internal/External"}),
    ReplaceInternalMCS = Table.ReplaceValue(ReplaceExternalEXT, "Internal MCS (DEPT)", "W/in Own Product Co", Replacer.ReplaceText, {"Internal/External"}),
    ReplaceInternalGE = Table.ReplaceValue(ReplaceInternalMCS, "Internal GE (OGE)", "W/in GE Outside BH", Replacer.ReplaceText, {"Internal/External"}),
    ReplaceInternalBH = Table.ReplaceValue(ReplaceInternalGE, "W/in BH not ProdCo", "W/in BH not ProdCo", Replacer.ReplaceText, {"Internal/External"}),
    
    // Standardize country names
    ReplaceUAE = Table.ReplaceValue(ReplaceInternalBH, "Utd.Arab.Emir.", "United Arab Emirates", Replacer.ReplaceText, {"Country"}),
    
    // Handle null Internal/External values
    ReplaceNullIntExt = Table.ReplaceValue(ReplaceUAE, null, "#", Replacer.ReplaceValue, {"Internal/External"}),
    ReplaceNullNotAssigned = Table.ReplaceValue(ReplaceNullIntExt, null, "Not assigned", Replacer.ReplaceValue, {"Internal/External"}),
    
    // Set final types
    SetFinalTypes = Table.TransformColumnTypes(ReplaceNullNotAssigned, {
        {"Company Code", type text}, 
        {"PROJECT TYPE", type text}, 
        {"WBS Person Responsible", type text}
    })
in
    SetFinalTypes


// ============================================================================
// QUERY: BO Sales 2020
// PURPOSE: Load 2020 Sales data from historical archive
// SOURCE: 2020 Sales Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20TS%20Sales%20-%20Histroical/Current%20Week/Previous%20Year-Dont%20Change/2020%20Sales.xlsx"), 
        null, 
        true
    ),
    
    SalesSheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(SalesSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Project Number", type text}, 
        {"Project Name", type text}, 
        {"Region", type text}, 
        {"Country", type text}, 
        {"Internal/External", type text}, 
        {"Sold To", type text}, 
        {"Sold to Description", type text}, 
        {"PM/SM Name", type text}, 
        {"Company Code", Int64.Type}, 
        {"Order Type", type text}, 
        {"Business", type text}, 
        {"Product Line", type text}, 
        {"PHx DS 17 PLs", type text}, 
        {"WBS NUMBER", type text}, 
        {"WBS DESCRIPTION", type text}, 
        {"WBS PRODUCT LINE", type text}, 
        {"MONTH", type text}, 
        {"QTR", Int64.Type}, 
        {"YEAR", Int64.Type}, 
        {"AMOUNT USD", type number}, 
        {"PROJECT TYPE", type text}, 
        {"WBS Person Responsible", type text}
    }),
    
    HandleProjectTypeErrors = Table.ReplaceErrorValues(SetColumnTypes, {{"PROJECT TYPE", "Not Assigned"}}),
    FilterValidProjects = Table.SelectRows(HandleProjectTypeErrors, each [Project Number] <> null and [Project Number] <> ""),
    ReplaceNullIntExt = Table.ReplaceValue(FilterValidProjects, null, "Not assigned", Replacer.ReplaceValue, {"Internal/External"}),
    SetCompanyCodeType = Table.TransformColumnTypes(ReplaceNullIntExt, {{"Company Code", type text}})
in
    SetCompanyCodeType


// ============================================================================
// QUERY: BO Sales 2021
// PURPOSE: Load 2021 Sales data from historical archive
// SOURCE: 2021 Sales Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20TS%20Sales%20-%20Histroical/Current%20Week/Previous%20Year-Dont%20Change/2021%20Sales.xlsx"), 
        null, 
        true
    ),
    
    SalesSheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(SalesSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Project Number", type text}, 
        {"Project Name", type text}, 
        {"Region", type text}, 
        {"Country", type text}, 
        {"Internal/External", type text}, 
        {"Sold To", type text}, 
        {"Sold to Description", type text}, 
        {"PM/SM Name", type text}, 
        {"Company Code", Int64.Type}, 
        {"Order Type", type text}, 
        {"Business", type text}, 
        {"Product Line", type text}, 
        {"PHx DS 17 PLs", type text}, 
        {"WBS NUMBER", type text}, 
        {"WBS DESCRIPTION", type text}, 
        {"WBS PRODUCT LINE", type text}, 
        {"MONTH", type text}, 
        {"QTR", Int64.Type}, 
        {"YEAR", Int64.Type}, 
        {"AMOUNT USD", type number}, 
        {"WBS Person Responsible", type text}, 
        {"PROJECT TYPE", type text}
    }),
    
    FilterValidProjects = Table.SelectRows(SetColumnTypes, each [Project Number] <> null and [Project Number] <> ""),
    ReplaceNullIntExt = Table.ReplaceValue(FilterValidProjects, null, "Not assigned", Replacer.ReplaceValue, {"Internal/External"}),
    SetCompanyCodeType = Table.TransformColumnTypes(ReplaceNullIntExt, {{"Company Code", type text}}),
    ReplaceNullProjectType = Table.ReplaceValue(SetCompanyCodeType, null, "Not Found", Replacer.ReplaceValue, {"PROJECT TYPE"}),
    HandleProjectTypeErrors = Table.ReplaceErrorValues(ReplaceNullProjectType, {{"PROJECT TYPE", "Not Found"}}),
    ReplaceNullResponsible = Table.ReplaceValue(HandleProjectTypeErrors, null, "Not Found", Replacer.ReplaceValue, {"WBS Person Responsible"}),
    HandleResponsibleErrors = Table.ReplaceErrorValues(ReplaceNullResponsible, {{"WBS Person Responsible", "Not Found"}})
in
    HandleResponsibleErrors


// ============================================================================
// QUERY: BO Sales 2022
// PURPOSE: Load 2022 Sales data from historical archive
// SOURCE: 2022 Sales Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20TS%20Sales%20-%20Histroical/Current%20Week/Previous%20Year-Dont%20Change/2022%20Sales.xlsx"), 
        null, 
        true
    ),
    
    SalesSheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(SalesSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Project Number", type text}, 
        {"Project Name", type text}, 
        {"Region", type text}, 
        {"Country", type text}, 
        {"Internal/External", type text}, 
        {"Sold To", type text}, 
        {"Sold to Description", type text}, 
        {"PM/SM Name", type text}, 
        {"Company Code", Int64.Type}, 
        {"Order Type", type text}, 
        {"Business", type text}, 
        {"Product Line", type text}, 
        {"PHx DS 17 PLs", type text}, 
        {"WBS NUMBER", type text}, 
        {"WBS DESCRIPTION", type text}, 
        {"WBS PRODUCT LINE", type text}, 
        {"MONTH", type text}, 
        {"QTR", Int64.Type}, 
        {"YEAR", Int64.Type}, 
        {"AMOUNT USD", type number}, 
        {"WBS Person Responsible", type text}, 
        {"PROJECT TYPE", type text}
    }),
    
    HandleSoldToErrors = Table.ReplaceErrorValues(SetColumnTypes, {{"Sold To", null}}),
    SetSoldToType = Table.TransformColumnTypes(HandleSoldToErrors, {{"Sold To", type text}}),
    FilterValidProjects = Table.SelectRows(SetSoldToType, each [Project Number] <> null and [Project Number] <> ""),
    ReplaceNullIntExt = Table.ReplaceValue(FilterValidProjects, null, "Not assigned", Replacer.ReplaceValue, {"Internal/External"}),
    SetCompanyCodeType = Table.TransformColumnTypes(ReplaceNullIntExt, {{"Company Code", type text}}),
    ReplaceNullResponsible = Table.ReplaceValue(SetCompanyCodeType, null, "Not Found", Replacer.ReplaceValue, {"WBS Person Responsible"}),
    HandleResponsibleErrors = Table.ReplaceErrorValues(ReplaceNullResponsible, {{"WBS Person Responsible", "Not Found"}}),
    ReplaceNullProjectType = Table.ReplaceValue(HandleResponsibleErrors, null, "Not Found", Replacer.ReplaceValue, {"PROJECT TYPE"}),
    HandleProjectTypeErrors = Table.ReplaceErrorValues(ReplaceNullProjectType, {{"PROJECT TYPE", "Not Found"}})
in
    HandleProjectTypeErrors


// ============================================================================
// QUERY: BO Sales 2023
// PURPOSE: Load 2023 Sales data from historical archive
// SOURCE: 2023 Sales Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20TS%20Sales%20-%20Histroical/Current%20Week/Previous%20Year-Dont%20Change/2023%20Sales.xlsx"), 
        null, 
        true
    ),
    
    SalesSheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(SalesSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Project Number", type text}, 
        {"Project Name", type text}, 
        {"Region", type text}, 
        {"Country", type text}, 
        {"Internal/External", type text}, 
        {"Sold To", type text}, 
        {"Sold to Description", type text}, 
        {"PM/SM Name", type text}, 
        {"Company Code", Int64.Type}, 
        {"Order Type", type text}, 
        {"Business", type text}, 
        {"Product Line", type text}, 
        {"PHx DS 17 PLs", type text}, 
        {"WBS NUMBER", type text}, 
        {"WBS DESCRIPTION", type text}, 
        {"WBS PRODUCT LINE", type text}, 
        {"MONTH", type text}, 
        {"QTR", Int64.Type}, 
        {"YEAR", Int64.Type}, 
        {"AMOUNT USD", type number}, 
        {"WBS Person Responsible", type text}, 
        {"PROJECT TYPE", type text}
    }),
    
    FilterValidProjects = Table.SelectRows(SetColumnTypes, each [Project Number] <> null and [Project Number] <> ""),
    HandleProjectTypeErrors = Table.ReplaceErrorValues(FilterValidProjects, {{"PROJECT TYPE", "Not Found"}}),
    ReplaceNullProjectType = Table.ReplaceValue(HandleProjectTypeErrors, null, "Not Found", Replacer.ReplaceValue, {"PROJECT TYPE"}),
    HandleResponsibleErrors = Table.ReplaceErrorValues(ReplaceNullProjectType, {{"WBS Person Responsible", "Not Found"}}),
    ReplaceNullResponsible = Table.ReplaceValue(HandleResponsibleErrors, null, "Not Found", Replacer.ReplaceValue, {"WBS Person Responsible"}),
    ReplaceNullIntExt = Table.ReplaceValue(ReplaceNullResponsible, null, "Not assigned", Replacer.ReplaceValue, {"Internal/External"}),
    SetCompanyCodeType = Table.TransformColumnTypes(ReplaceNullIntExt, {{"Company Code", type text}})
in
    SetCompanyCodeType


// ============================================================================
// QUERY: TS_CY_Sales (Current Year Sales)
// PURPOSE: Load current year (2025) Sales data from live source
// SOURCE: CQ Sales_Backlog Excel file - CY Sales sheet
// BUSINESS USE: Real-time current year sales tracking
// ============================================================================
let
    // Connect to CQ Sales/Backlog file
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/CQ%20Sales%20and%20Backlog%20(Projects%20and%20Services%20Pacing%20backup)/CQ%20Sales%20_Backlog.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to CY Sales sheet
    CYSalesSheet = Source{[Item="CY Sales",Kind="Sheet"]}[Data],
    
    // Promote headers
    PromotedHeaders = Table.PromoteHeaders(CYSalesSheet, [PromoteAllScalars=true]),
    
    // Remove system columns and reorder
    RemoveSystemColumns = Table.RemoveColumns(PromotedHeaders, {"BW Status: SYS0", "Sales row number"}),
    
    // Duplicate columns for processing
    DuplicateProductLine = Table.DuplicateColumn(RemoveSystemColumns, "WBS Product Line", "wbs prod line - Copy"),
    DuplicateSnapshotDate = Table.DuplicateColumn(DuplicateProductLine, "Snapshot Date", "Snapshot Date - Copy"),
    RenameRefreshDate = Table.RenameColumns(DuplicateSnapshotDate, {{"Snapshot Date - Copy", "Refresh Date"}}),
    
    // Split WBS Product Line to extract key
    SplitProductLine = Table.SplitColumn(RenameRefreshDate, "wbs prod line - Copy", Splitter.SplitTextByEachDelimiter({"-"}, QuoteStyle.Csv, false), {"wbs prod line - Copy.1", "wbs prod line - Copy.2"}),
    
    SetProductLineKeyType = Table.TransformColumnTypes(SplitProductLine, {{"wbs prod line - Copy.1", Int64.Type}, {"wbs prod line - Copy.2", type text}}),
    RenameProductLineKey = Table.RenameColumns(SetProductLineKeyType, {{"wbs prod line - Copy.1", "WBS Prod Line Key"}}),
    RemoveProductLineSuffix = Table.RemoveColumns(RenameProductLineKey, {"wbs prod line - Copy.2"}),
    SetProductLineKeyText = Table.TransformColumnTypes(RemoveProductLineSuffix, {{"WBS Prod Line Key", type text}}),
    
    // Rename columns to standard names
    RenameRegion = Table.RenameColumns(SetProductLineKeyText, {{"Region", "Region-Delete"}}),
    RenameProject = Table.RenameColumns(RenameRegion, {{"Project Definition", "Project Number"}}),
    RenameProjectName = Table.RenameColumns(RenameProject, {{"Project Description", "Project Name"}}),
    RenameWBSNumber = Table.RenameColumns(RenameProjectName, {{"WBS Element Number", "WBS NUMBER"}}),
    RenameWBSDescription = Table.RenameColumns(RenameWBSNumber, {{"WBS Element Description", "WBS DESCRIPTION"}}),
    RenameIntExt = Table.RenameColumns(RenameWBSDescription, {{"Internal/External Text", "Internal/External"}}),
    RenameSoldToDesc = Table.RenameColumns(RenameIntExt, {{"Customer Sold To Name", "Sold to Description"}}),
    RenameSalesValue = Table.RenameColumns(RenameSoldToDesc, {{"Total Sales Value USD", "AMOUNT USD"}}),
    RenameCountry = Table.RenameColumns(RenameSalesValue, {{"Sub region", "Country"}}),
    RenameCompanyCode = Table.RenameColumns(RenameCountry, {{"Company Code Name", "Company Code"}}),
    
    // Remove additional unused columns
    RemoveUnusedColumns = Table.RemoveColumns(RenameCompanyCode, {
        "Customer Type", "Backlog row number", "Company Code Key", 
        "Milestone Incremental POC", "WBS Priority"
    }),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(RemoveUnusedColumns, {
        {"Snapshot Date", type date}, 
        {"Sales Order Currency", type text}
    }),
    
    RemovePoleRegionKey = Table.RemoveColumns(SetColumnTypes, {"Pole Region Key"}),
    
    // Rename and process Sales Order columns
    RenameSoldTo = Table.RenameColumns(RemovePoleRegionKey, {{"Customer Sold To Number", "Sold To"}}),
    RenameSONumber = Table.RenameColumns(RenameSoldTo, {{"Sales Order Number", "SO Number"}}),
    SetSONumberType = Table.TransformColumnTypes(RenameSONumber, {{"SO Number", type text}}),
    
    // Rename WBS columns
    RenameWBSProductLine = Table.RenameColumns(SetSONumberType, {{"WBS Product Line", "WBS PRODUCT LINE"}}),
    RenameResponsiblePerson = Table.RenameColumns(RenameWBSProductLine, {{"WBS Responsible Person Text", "PM/SM Name"}}),
    RenameOrderType = Table.RenameColumns(RenameResponsiblePerson, {{"Sales Order Type Key", "SO Type"}}),
    
    // Process date columns
    DuplicateSalesDate = Table.DuplicateColumn(RenameOrderType, "Calculated Sales Date", "Calculated Sales Date - Copy"),
    ExtractYear = Table.TransformColumns(DuplicateSalesDate, {{"Calculated Sales Date - Copy", each Number.IntegerDivide(_, 10000), Int64.Type}}),
    RenameYear = Table.RenameColumns(ExtractYear, {{"Calculated Sales Date - Copy", "YEAR"}}),
    
    DuplicateForMonth = Table.DuplicateColumn(RenameYear, "Calculated Sales Date", "Calculated Sales Date - Copy"),
    ExtractMonthRaw = Table.TransformColumns(DuplicateForMonth, {{"Calculated Sales Date - Copy", each Number.Mod(Number.IntegerDivide(_, 100), 100), Int64.Type}}),
    RenameMonthRaw = Table.RenameColumns(ExtractMonthRaw, {{"Calculated Sales Date - Copy", "Sales Month"}}),
    
    // Convert month number to month name
    AddMonthName = Table.AddColumn(RenameMonthRaw, "MonthName", each 
        if [Sales Month] = 1 then "JAN"
        else if [Sales Month] = 2 then "FEB"
        else if [Sales Month] = 3 then "MAR"
        else if [Sales Month] = 4 then "APR"
        else if [Sales Month] = 5 then "MAY"
        else if [Sales Month] = 6 then "JUN"
        else if [Sales Month] = 7 then "JUL"
        else if [Sales Month] = 8 then "AUG"
        else if [Sales Month] = 9 then "SEP"
        else if [Sales Month] = 10 then "OCT"
        else if [Sales Month] = 11 then "NOV"
        else if [Sales Month] = 12 then "DEC"
        else null
    ),
    
    // Extract Quarter
    AddQuarter = Table.AddColumn(AddMonthName, "QTR", each Number.RoundUp([Sales Month] / 3)),
    SetQuarterType = Table.TransformColumnTypes(AddQuarter, {{"QTR", type text}}),
    
    // Remove and rename columns
    RemoveSalesMonth = Table.RemoveColumns(SetQuarterType, {"Sales Month"}),
    RenameMONTH = Table.RenameColumns(RemoveSalesMonth, {{"MonthName", "MONTH"}}),
    
    // Remove additional unused columns
    RemoveMoreColumns = Table.RemoveColumns(RenameMONTH, {
        "Milestone Description", "Milestone Usage", "Overdue Revenue Flag",
        "PH/WBS Product Line Key", "PH3 Family Text", "PH4 Sub Line Text",
        "PH5 Model Text", "Project / Flow Indicator", "Sales Order Sales Office Text",
        "Sales Order Version Number", "WBS MLST Forecast Fixed Date (Rev Rec)",
        "Project Responsible Person Number", "Calculated Sales Date",
        "Milestone Number", "Calculated Sales Year Quarter", "WBS RA Key",
        "Customer PO Number", "PH2 Product Line Text", "Sales Year",
        "WBS Functional Area", "WBS Project Type", "Sales Order Created On Date",
        "Sales Order Created By text", "SO Line Number", "Sales Order Currency",
        "WBS Status Text", "Snapshot Flag", "Snapshot Date", "WBS Prod Line Key",
        "Refresh Date", "Project Responsible Person Text"
    }),
    
    // Add PROJECT TYPE column (not available in this source)
    AddProjectType = Table.AddColumn(RemoveMoreColumns, "PROJECT TYPE", each "Not Found"),
    SetProjectTypeType = Table.TransformColumnTypes(AddProjectType, {{"PROJECT TYPE", type text}}),
    
    // Rename End User column
    RenameEndUser = Table.RenameColumns(SetProjectTypeType, {{"End User Name", "End user_Ex"}}),
    SetEndUserType = Table.TransformColumnTypes(RenameEndUser, {{"End user_Ex", type text}}),
    
    RemoveRegionNew = Table.RemoveColumns(SetEndUserType, {"Region New"}),
    
    // Rename WBS responsible person
    RenameWBSResponsible = Table.RenameColumns(RemoveRegionNew, {{"WBS Person Resp", "WBS Person Responsible"}}),
    
    // Final column cleanup
    RemoveFinalColumns = Table.RemoveColumns(RenameWBSResponsible, {"SO Type", "wbs prod line - Copy", "MonthName", "Month No."}),
    SetWBSResponsibleType = Table.TransformColumnTypes(RemoveFinalColumns, {{"WBS Person Responsible", type text}}),
    
    // Remove "Q" prefix from Quarter
    CleanQuarter = Table.ReplaceValue(SetWBSResponsibleType, "Q", "", Replacer.ReplaceText, {"QTR"}),
    
    // Add PH product line column
    DuplicateProductLineForPH = Table.DuplicateColumn(CleanQuarter, "Product Line", "Product Line - Copy"),
    RenamePHColumn = Table.RenameColumns(DuplicateProductLineForPH, {{"Product Line - Copy", "PHx DS 17 PLs"}}),
    
    // Reorder columns to match historical format
    ReorderColumns = Table.ReorderColumns(RenamePHColumn, {
        "Project Number", "Project Name", "Region", "Country", "Internal/External", 
        "Sold To", "Sold to Description", "PM/SM Name", "Order Type", "Product Line", 
        "PHx DS 17 PLs", "WBS NUMBER", "WBS DESCRIPTION", "WBS PRODUCT LINE", 
        "MONTH", "QTR", "YEAR", "AMOUNT USD", "WBS Person Responsible"
    }),
    
    // Standardize Product Line names to PH codes
    // Business Rule: Map legacy product line names to standard PH2 codes
    ReplaceBNSW = Table.ReplaceValue(ReorderColumns, "BN Software", "BNSW", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSSV = Table.ReplaceValue(ReplaceBNSW, "BN SW Service", "BSSV", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBNHW = Table.ReplaceValue(ReplaceBSSV, "BN Hardware", "BNHW", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBHSV = Table.ReplaceValue(ReplaceBNHW, "BN HW Service", "BHSV", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSSH = Table.ReplaceValue(ReplaceBHSV, "BN SW SVS - Host", "BSSH", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSSA = Table.ReplaceValue(ReplaceBSSH, "BN SW SVS - ARMS", "BSSA", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSSG = Table.ReplaceValue(ReplaceBSSA, "BN SW SVS - Augury", "BSSG", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSHG = Table.ReplaceValue(ReplaceBSSG, "BN HW SVS - Augury", "BSHG", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSHH = Table.ReplaceValue(ReplaceBSHG, "BN HW SVS - Host", "BSHH", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSSX = Table.ReplaceValue(ReplaceBSHH, "BN SW SVS - VitalyX", "BSSX", Replacer.ReplaceText, {"Product Line"}),
    ReplaceBSHX = Table.ReplaceValue(ReplaceBSSX, "BN HW SVS - VitalyX", "BSHX", Replacer.ReplaceText, {"Product Line"})
in
    ReplaceBSHX


// ============================================================================
// QUERY: BOSales20202025
// PURPOSE: Combine all historical and current year sales data
// SOURCE: Appended from BO Sales 2020-2024 and TS_CY_Sales
// BUSINESS USE: Unified sales dataset for multi-year analysis
// ============================================================================
let
    // Combine all sales year tables
    Source = Table.Combine({#"BO Sales 2024", #"BO Sales 2020", #"BO Sales 2021", #"BO Sales 2022", #"BO Sales 2023", TS_CY_Sales}),
    
    // Handle Internal/External classification edge cases
    // Business Rule: If Sold To is "#" (placeholder), classify as External
    AddIntExtNormalized = Table.AddColumn(Source, "Internal/External_N", each 
        if [Sold To] = "#" then "External Customers" 
        else [#"Internal/External"]
    ),
    
    // Reorder and clean up columns
    ReorderColumns = Table.ReorderColumns(AddIntExtNormalized, {
        "Project Number", "Project Name", "Region", "Country", "Internal/External", 
        "Internal/External_N", "Sold To", "Sold to Description", "PM/SM Name", 
        "Company Code", "Order Type", "Business", "Product Line", "PHx DS 17 PLs", 
        "WBS NUMBER", "WBS DESCRIPTION", "WBS PRODUCT LINE", "MONTH", "QTR", 
        "YEAR", "AMOUNT USD"
    }),
    
    // Remove original Internal/External, rename normalized version
    RemoveOldIntExt = Table.RemoveColumns(ReorderColumns, {"Internal/External"}),
    RenameIntExt = Table.RenameColumns(RemoveOldIntExt, {{"Internal/External_N", "Internal/External"}}),
    
    // Handle null amounts
    ReplaceNullAmount = Table.ReplaceValue(RenameIntExt, null, 0, Replacer.ReplaceValue, {"AMOUNT USD"}),
    
    // Filter valid projects
    FilterValidProjects = Table.SelectRows(ReplaceNullAmount, each [Project Number] <> null and [Project Number] <> ""),
    
    // Handle null Internal/External
    ReplaceNullIntExt = Table.ReplaceValue(FilterValidProjects, null, "Not assigned", Replacer.ReplaceValue, {"Internal/External"}),
    
    // Standardize country names (case normalization)
    ReplaceIreland = Table.ReplaceValue(ReplaceNullIntExt, "IRELAND", "Ireland", Replacer.ReplaceText, {"Country"}),
    ReplaceTunisia = Table.ReplaceValue(ReplaceIreland, "TUNISIA", "Tunisia", Replacer.ReplaceText, {"Country"}),
    ReplaceLibya = Table.ReplaceValue(ReplaceTunisia, "LIBYA", "Libya", Replacer.ReplaceText, {"Country"}),
    ReplaceAndorra = Table.ReplaceValue(ReplaceLibya, "ANDORRA", "Andorra", Replacer.ReplaceText, {"Country"}),
    
    // Set column types
    SetColumnTypes = Table.TransformColumnTypes(ReplaceAndorra, {
        {"Internal/External", type text}, 
        {"QTR", type text}, 
        {"YEAR", type text}
    }),
    
    // Handle edge case Order Types
    ReplaceOrderType0 = Table.ReplaceValue(SetColumnTypes, "0", "#", Replacer.ReplaceText, {"Order Type"}),
    ReplaceNullOrderType = Table.ReplaceValue(ReplaceOrderType0, null, "#", Replacer.ReplaceValue, {"Order Type"}),
    
    // Handle edge case Internal/External values
    ReplaceIntExt0 = Table.ReplaceValue(ReplaceNullOrderType, "0", "#", Replacer.ReplaceText, {"Internal/External"}),
    ReplaceDoNotUse = Table.ReplaceValue(ReplaceIntExt0, "Do Not Use", "#", Replacer.ReplaceText, {"Internal/External"}),
    
    // Set End User type
    SetEndUserType = Table.TransformColumnTypes(ReplaceDoNotUse, {{"End user_Ex", type text}})
in
    SetEndUserType


// ============================================================================
// QUERY: Internal External Standard Map_Sales
// PURPOSE: Internal/External classification mapping for Sales reports
// SOURCE: SMF Mapping File Master - Internal External Standard Map sheet
// BUSINESS USE: Standardizes customer type classification for Sales reporting
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    MappingSheet = Source{[Item="Internal External Standard Map",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(MappingSheet, [PromoteAllScalars=true]),
    
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
    
    // Filter valid rows
    FilterValidStandard = Table.SelectRows(SetColumnTypes, each [#"Standard Internal/External "] <> null and [#"Standard Internal/External "] <> ""),
    
    // Remove columns not needed for Sales report mapping
    RemoveOtherReportColumns = Table.RemoveColumns(FilterValidStandard, {
        "Order Report", "Backlog Report", "Convertible Order", 
        "As Sold Margin", "As Sold Report", "PLM", "Unbilled BDR", "Sales WORP"
    }),
    
    // Filter to rows with Sales Report mapping
    FilterValidSalesReport = Table.SelectRows(RemoveOtherReportColumns, each [Sales Report] <> null and [Sales Report] <> "")
in
    FilterValidSalesReport


// ============================================================================
// QUERY: BO Sales Data_Summarize
// PURPOSE: Summarized Sales data for Services product lines (excluding current quarter)
// SOURCE: BO Sales Data
// BUSINESS USE: Historical services revenue analysis (excludes in-progress quarter)
// ============================================================================
let
    Source = #"BO Sales Data",
    
    // Filter to Services product lines only
    FilterServices = Table.SelectRows(Source, each ([Services PL Flag] = "Services")),
    
    // Exclude current quarter (in-progress data)
    FilterHistorical = Table.SelectRows(FilterServices, each [Current Quarter] = "No"),
    
    // Handle errors and nulls
    HandleSalesErrors = Table.ReplaceErrorValues(FilterHistorical, {{"Sales Amount", 0}}),
    ReplaceNullSales = Table.ReplaceValue(HandleSalesErrors, null, 0, Replacer.ReplaceValue, {"Sales Amount"}),
    ReplaceNullCountry = Table.ReplaceValue(ReplaceNullSales, null, "Not assigned", Replacer.ReplaceValue, {"Standard Country"}),
    ReplaceNullIntExt = Table.ReplaceValue(ReplaceNullCountry, null, "Not assigned", Replacer.ReplaceValue, {"INT/EXT Type"}),
    
    SetColumnTypes = Table.TransformColumnTypes(ReplaceNullIntExt, {
        {"INT/EXT Type", type text}, 
        {"Standard Country", type text}
    }),
    
    // Group by region and time dimensions
    GroupedData = Table.Group(SetColumnTypes, {"Standard Region", "Standard Sub Region", "Standard Country", "YearMonthDate"}, {
        {"Total Sales Amount USD", each List.Sum([Sales Amount]), type nullable number}
    }),
    
    // Rename for consistency
    RenameValueColumn = Table.RenameColumns(GroupedData, {{"Total Sales Amount USD", "Value"}}),
    SetValueType = Table.TransformColumnTypes(RenameValueColumn, {{"Value", type number}})
in
    SetValueType


// ============================================================================
// QUERY: BO Sales Data_Summarize_For NAM
// PURPOSE: Summarized Sales data for NAM Services (excluding current quarter)
// SOURCE: BO Sales Data
// BUSINESS USE: NAM regional services revenue for CM analysis
// ============================================================================
let
    Source = #"BO Sales Data",
    
    // Filter to Services only
    FilterServices = Table.SelectRows(Source, each ([Services PL Flag] = "Services")),
    
    // Exclude current quarter
    FilterHistorical = Table.SelectRows(FilterServices, each [Current Quarter] = "No"),
    
    // Filter to NAM region only
    FilterNAM = Table.SelectRows(FilterHistorical, each ([Standard Region] = "NAM")),
    
    // Handle errors and nulls
    HandleSalesErrors = Table.ReplaceErrorValues(FilterNAM, {{"Sales Amount", 0}}),
    ReplaceNullSales = Table.ReplaceValue(HandleSalesErrors, null, 0, Replacer.ReplaceValue, {"Sales Amount"}),
    ReplaceNullCountry = Table.ReplaceValue(ReplaceNullSales, null, "Not assigned", Replacer.ReplaceValue, {"Standard Country"}),
    ReplaceNullIntExt = Table.ReplaceValue(ReplaceNullCountry, null, "Not assigned", Replacer.ReplaceValue, {"INT/EXT Type"}),
    
    SetColumnTypes = Table.TransformColumnTypes(ReplaceNullIntExt, {
        {"INT/EXT Type", type text}, 
        {"Standard Country", type text}
    }),
    
    // Group by sub-region for NAM detail
    GroupedData = Table.Group(SetColumnTypes, {"Standard Region", "Standard Sub Region", "Standard Country", "YearMonthDate"}, {
        {"Total Sales Amount USD", each List.Sum([Sales Amount]), type nullable number}
    }),
    
    RenameValueColumn = Table.RenameColumns(GroupedData, {{"Total Sales Amount USD", "Value"}}),
    SetValueType = Table.TransformColumnTypes(RenameValueColumn, {{"Value", type number}})
in
    SetValueType


// ============================================================================
// QUERY: BO Sales Data_Summarize_For NAM_total Revenue
// PURPOSE: Total Revenue (all product lines) for NAM region
// SOURCE: BO Sales Data
// BUSINESS USE: Complete NAM revenue for contribution margin calculation
// ============================================================================
let
    Source = #"BO Sales Data",
    
    // No filter on Services - include all product lines
    // Exclude current quarter only
    FilterHistorical = Table.SelectRows(Source, each [Current Quarter] = "No"),
    
    // Filter to NAM only
    FilterNAM = Table.SelectRows(FilterHistorical, each ([Standard Region] = "NAM")),
    
    // Handle errors and nulls
    HandleSalesErrors = Table.ReplaceErrorValues(FilterNAM, {{"Sales Amount", 0}}),
    ReplaceNullSales = Table.ReplaceValue(HandleSalesErrors, null, 0, Replacer.ReplaceValue, {"Sales Amount"}),
    ReplaceNullCountry = Table.ReplaceValue(ReplaceNullSales, null, "Not assigned", Replacer.ReplaceValue, {"Standard Country"}),
    ReplaceNullIntExt = Table.ReplaceValue(ReplaceNullCountry, null, "Not assigned", Replacer.ReplaceValue, {"INT/EXT Type"}),
    
    SetColumnTypes = Table.TransformColumnTypes(ReplaceNullIntExt, {
        {"INT/EXT Type", type text}, 
        {"Standard Country", type text}
    }),
    
    // Group by sub-region
    GroupedData = Table.Group(SetColumnTypes, {"Standard Region", "Standard Sub Region", "Standard Country", "YearMonthDate"}, {
        {"Total Revenue", each List.Sum([Sales Amount]), type nullable number}
    }),
    
    SetTotalRevenueType = Table.TransformColumnTypes(GroupedData, {{"Total Revenue", type number}})
in
    SetTotalRevenueType


// ============================================================================
// QUERY: BO Sales Data (Main Transformed Sales Fact Table)
// PURPOSE: Master sales fact table with full transformations and enrichments
// SOURCE: BOSales20202025 combined data
// BUSINESS USE: Primary sales analysis dataset with geographic standardization,
//               product hierarchy, and business classification
// DEPENDENCIES: BOSales20202025, Standard Country List -Sales, Product Lines Map,
//               Internal External Standard Map_Sales, EndUserMapTBL
// ============================================================================
let
    // Start from combined sales data
    Source = BOSales20202025,
    
    // Ensure Sold To is text type
    SetSoldToType = Table.TransformColumnTypes(Source, {{"Sold To", type text}}),
    
    // Set all base column types
    SetBaseColumnTypes = Table.TransformColumnTypes(SetSoldToType, {
        {"Project Number", type text}, 
        {"Project Name", type text}, 
        {"Region", type text}, 
        {"Country", type text}, 
        {"Internal/External", type text}, 
        {"Sold To", type text}, 
        {"Sold to Description", type text}, 
        {"PM/SM Name", type text}, 
        {"Company Code", type text}, 
        {"Order Type", type text}, 
        {"Business", type text}, 
        {"Product Line", type text}, 
        {"PHx DS 17 PLs", type text}, 
        {"WBS NUMBER", type text}, 
        {"WBS DESCRIPTION", type text}, 
        {"WBS PRODUCT LINE", type text}, 
        {"MONTH", type text}, 
        {"QTR", Int64.Type}, 
        {"YEAR", Int64.Type}, 
        {"AMOUNT USD", type number}
    }),
    
    // Rename columns to standard names
    RenameColumns = Table.RenameColumns(SetBaseColumnTypes, {
        {"Region", "Region_Delete"}, 
        {"Country", "Country_Delete"}, 
        {"PM/SM Name", "Responsible Person"}, 
        {"WBS NUMBER", "WBS Number"}, 
        {"WBS PRODUCT LINE", "WBS Product Line"}, 
        {"WBS DESCRIPTION", "WBS Description"}, 
        {"MONTH", "Month"}, 
        {"QTR", "Quarter"}, 
        {"YEAR", "Year"}
    }),
    
    // Convert month name to month number
    // Business Rule: Standard month number format for date calculations
    AddMonthNumber = Table.AddColumn(RenameColumns, "Month New", each 
        if [Month] = "JAN" then 1 
        else if [Month] = "FEB" then 2 
        else if [Month] = "MAR" then 3 
        else if [Month] = "APR" then 4 
        else if [Month] = "MAY" then 5 
        else if [Month] = "JUN" then 6 
        else if [Month] = "JUL" then 7 
        else if [Month] = "AUG" then 8 
        else if [Month] = "SEP" then 9 
        else if [Month] = "OCT" then 10 
        else if [Month] = "NOV" then 11 
        else if [Month] = "DEC" then 12 
        else [Month]
    ),
    
    // Standardize country names (case corrections)
    ReplaceIreland = Table.ReplaceValue(AddMonthNumber, "IRELAND", "Ireland", Replacer.ReplaceText, {"Country_Delete"}),
    ReplaceTunisia = Table.ReplaceValue(ReplaceIreland, "TUNISIA", "Tunisia", Replacer.ReplaceText, {"Country_Delete"}),
    ReplaceLibya = Table.ReplaceValue(ReplaceTunisia, "LIBYA", "Libya", Replacer.ReplaceText, {"Country_Delete"}),
    ReplaceAndorra = Table.ReplaceValue(ReplaceLibya, "ANDORRA", "Andorra", Replacer.ReplaceText, {"Country_Delete"}),
    
    // Reorder columns
    ReorderColumns = Table.ReorderColumns(ReplaceAndorra, {
        "Project Number", "Project Name", "Region_Delete", "Country_Delete", 
        "Internal/External", "Sold To", "Sold to Description", "Responsible Person", 
        "Company Code", "Order Type", "Business", "Product Line", "PHx DS 17 PLs", 
        "WBS Number", "WBS Description", "WBS Product Line", "Month", "Month New", 
        "Quarter", "Year", "AMOUNT USD"
    }),
    
    // Handle null values in source data
    ReplaceNullCountry = Table.ReplaceValue(ReorderColumns, null, "Not assigned", Replacer.ReplaceValue, {"Country_Delete"}),
    ReplaceNullRegion = Table.ReplaceValue(ReplaceNullCountry, null, "Not assigned", Replacer.ReplaceValue, {"Region_Delete"}),
    
    // Remove original Month, use Month New
    RemoveOldMonth = Table.RemoveColumns(ReplaceNullRegion, {"Month"}),
    ReplaceNullWBSDesc = Table.ReplaceValue(RemoveOldMonth, null, "Not Assigned", Replacer.ReplaceValue, {"WBS Description"}),
    RenameMonthNew = Table.RenameColumns(ReplaceNullWBSDesc, {{"Month New", "Month"}}),
    ReplaceNullProjectName = Table.ReplaceValue(RenameMonthNew, null, "Not Assigned", Replacer.ReplaceValue, {"Project Name"}),
    RenameSalesAmount = Table.RenameColumns(ReplaceNullProjectName, {{"AMOUNT USD", "Total Sales Amount USD"}}),
    ReplaceHashCountry = Table.ReplaceValue(RenameSalesAmount, "#", "Not assigned", Replacer.ReplaceText, {"Country_Delete"}),
    
    // Add Function Flag for product categorization
    // Business Rule: Categorize by Product Line to Hardware/Software/Services
    AddFunctionFlag = Table.AddColumn(ReplaceHashCountry, "Function Flag", each 
        if [Product Line] = "BNHW" then "Hardware" 
        else if [Product Line] = "BHSW" then "Software" 
        else if [Product Line] = "BHHW" then "Hardware" 
        else if [Product Line] = "BNSW" then "Software" 
        else "Services"
    ),
    SetFunctionFlagType = Table.TransformColumnTypes(AddFunctionFlag, {{"Function Flag", type text}}),
    
    // Standardize region codes
    ReplaceEuropeRegion = Table.ReplaceValue(SetFunctionFlagType, "Europe", "EU", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceMENATRegion = Table.ReplaceValue(ReplaceEuropeRegion, "MENAT", "MEN", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceNARegion = Table.ReplaceValue(ReplaceMENATRegion, "North America", "NA", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceLARegion = Table.ReplaceValue(ReplaceNARegion, "Latin America", "LA", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceAsiaRegion = Table.ReplaceValue(ReplaceLARegion, "Asia", "AS", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceIndiaRegion = Table.ReplaceValue(ReplaceAsiaRegion, "India", "IN", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceCISRegion = Table.ReplaceValue(ReplaceIndiaRegion, "RUSSIA / CIS", "CIS", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceChinaRegion = Table.ReplaceValue(ReplaceCISRegion, "China", "CH", Replacer.ReplaceText, {"Region_Delete"}),
    ReplaceNullCountryFinal = Table.ReplaceValue(ReplaceChinaRegion, null, "Not assigned", Replacer.ReplaceValue, {"Country_Delete"}),
    
    // Apply geographic override rules
    // Business Rule: Mexico projects (MX- prefix) in NA should map to Latin America
    // Business Rule: US state prefix projects (PA-, NC-, TX-, NV-) for 1BN Services stay in NA
    AddRegionOverride = Table.AddColumn(ReplaceNullCountryFinal, "Region_Delete1", each 
        if Text.Contains([Project Name], "MX-") and [Region_Delete] = "NA" then "LA"
        else if (Text.StartsWith([Project Name], "PA-") or Text.StartsWith([Project Name], "NC-") or 
                 Text.StartsWith([Project Name], "TX-") or Text.StartsWith([Project Name], "NV-")) and 
                Text.StartsWith([Project Number], "1BN") and Text.Contains([Function Flag], "Services") and 
                Text.Contains([Region_Delete], "NA") then "NA"
        else if (Text.StartsWith([WBS Description], "PA-") or Text.StartsWith([WBS Description], "NC-") or 
                 Text.StartsWith([WBS Description], "TX-") or Text.StartsWith([WBS Description], "NV-")) and 
                Text.StartsWith([Project Number], "1BN") and Text.Contains([Function Flag], "Services") and 
                Text.Contains([Region_Delete], "NA") then "NA"
        else [Region_Delete]
    ),
    HandleRegionErrors = Table.ReplaceErrorValues(AddRegionOverride, {{"Region_Delete1", "error"}}),
    
    // Apply country override rules
    // Business Rule: 1BN Canada projects map to United States
    // Business Rule: US state prefix projects map to United States
    AddCountryOverride = Table.AddColumn(HandleRegionErrors, "Country_Delete1", each 
        if Text.Contains([Project Name], "MX-") and [Region_Delete] = "NA" then "Mexico"
        else if Text.Contains([Project Number], "1BN") and [Country_Delete] = "Canada" and 
                Text.Contains([Region_Delete], "NA") then "United States"
        else if (Text.StartsWith([Project Name], "PA-") or Text.StartsWith([Project Name], "NC-") or 
                 Text.StartsWith([Project Name], "TX-") or Text.StartsWith([Project Name], "NV-")) and 
                Text.StartsWith([Project Number], "1BN") and Text.Contains([Function Flag], "Services") and 
                Text.Contains([Region_Delete], "NA") then "United States"
        else if (Text.StartsWith([WBS Description], "PA-") or Text.StartsWith([WBS Description], "NC-") or 
                 Text.StartsWith([WBS Description], "TX-") or Text.StartsWith([WBS Description], "NV-")) and 
                Text.StartsWith([WBS Description], "1BN") and Text.Contains([Function Flag], "Services") and 
                Text.Contains([Region_Delete], "NA") then "United States"
        else [Country_Delete]
    ),
    HandleCountryErrors = Table.ReplaceErrorValues(AddCountryOverride, {{"Country_Delete1", "error"}}),
    
    // Clean up override columns
    RemoveOldRegionCountry = Table.RemoveColumns(HandleCountryErrors, {"Region_Delete", "Country_Delete"}),
    RenameOverrideColumns = Table.RenameColumns(RemoveOldRegionCountry, {
        {"Region_Delete1", "Region_Delete"}, 
        {"Country_Delete1", "Country_Delete"}
    }),
    SetOverrideTypes = Table.TransformColumnTypes(RenameOverrideColumns, {
        {"Region_Delete", type text}, 
        {"Country_Delete", type text}
    }),
    
    // Join with Standard Country List for geographic hierarchy
    JoinCountryList = Table.NestedJoin(SetOverrideTypes, {"Region_Delete", "Country_Delete"}, #"Standard Country List -Sales", {"Region-Sales Report", "Country-Sales Report"}, "Standard Country List", JoinKind.LeftOuter),
    ExpandCountryList = Table.ExpandTableColumn(JoinCountryList, "Standard Country List", {"Standard Country", "Standard Region", "Standard Sub Region"}, {"Standard Country", "Standard Region", "Standard Sub Region"}),
    
    // Join with Product Lines Map for product hierarchy
    JoinProductLines = Table.NestedJoin(ExpandCountryList, {"WBS Product Line"}, #"Product Lines Map", {"WBS Product Line"}, "Product Lines Map", JoinKind.LeftOuter),
    ExpandProductLines = Table.ExpandTableColumn(JoinProductLines, "Product Lines Map", {
        "PH1", "PH2", "PH3", "PH4", "PH5", "Revenue Stream", "Delivery Mode", 
        "SMF Mapping 1", "SMF Mapping 2", "Core Mapping"
    }, {
        "PH1", "PH2", "PH3", "PH4", "PH5", "Revenue Stream", "Delivery Mode", 
        "SMF Mapping 1", "SMF Mapping 2", "Core Mapping"
    }),
    
    // Add Order Type Category
    // Business Rule: ZSBO and L2 are Reactive orders, ZPFM is Backlog
    AddOrderTypeCategory = Table.AddColumn(ExpandProductLines, "Order Type Category", each 
        if [Order Type] = "ZSBO" then "Reactive" 
        else if [Order Type] = "L2" then "Reactive" 
        else if [Order Type] = "ZPFM" then "Backlog" 
        else "Not Classified"
    ),
    
    // Rename amount column and add Data Source identifier
    RenameTotalAmount = Table.RenameColumns(AddOrderTypeCategory, {{"Total Sales Amount USD", "Total Amount USD"}}),
    AddDataSource = Table.AddColumn(RenameTotalAmount, "Data Source", each 1),
    SetDataSourceType = Table.TransformColumnTypes(AddDataSource, {{"Data Source", type text}}),
    ReplaceDataSource = Table.ReplaceValue(SetDataSourceType, "1", "BO Sales", Replacer.ReplaceText, {"Data Source"}),
    
    // Add INT/EXT Type classification
    // Business Rule: External Customers, W/in GE Outside BH, W/in BH not ProdCo are External
    // Business Rule: W/in Own Product Co, Internal OG (GRP), Internal MCS (DEPT) are Internal
    AddIntExtType = Table.AddColumn(ReplaceDataSource, "INT/EXT Type", each 
        if [#"Internal/External"] = "External Customers" then "External" 
        else if [#"Internal/External"] = "W/in GE Outside BH" then "External" 
        else if [#"Internal/External"] = "W/in BH not ProdCo" then "External" 
        else if [#"Internal/External"] = "W/in Own Product Co" then "Internal" 
        else if [#"Internal/External"] = "Internal OG (GRP)" then "Internal" 
        else if [#"Internal/External"] = "Internal MCS (DEPT)" then "Internal" 
        else if [#"Internal/External"] = "Internal GE (OGE)" then "External" 
        else if [#"Internal/External"] = "External (EXT)" then "External" 
        else "Not assigned"
    ),
    
    // Add Services PL Flag based on PH2
    AddServicesPLFlag = Table.AddColumn(AddIntExtType, "Services PL Flag", each 
        if [PH2] = "BNHW" then "Hardware" 
        else if [PH2] = "BHSW" then "Software" 
        else if [PH2] = "BHHW" then "Hardware" 
        else if [PH2] = "BNSW" then "Software" 
        else "Services"
    ),
    
    // Add US Sub-Region breakdown for Services
    // Business Rule: US state prefixes (PA-, NC-, TX-, NV-) determine sub-region
    AddCountryEnhanced = Table.AddColumn(AddServicesPLFlag, "Country_en", each 
        if Text.Contains([Project Name], "PA-") and Text.Contains([Standard Country], "United States") then "US -North"
        else if Text.Contains([Project Name], "NC-") and Text.Contains([Standard Country], "United States") then "US -South"
        else if Text.Contains([Project Name], "TX-") and Text.Contains([Standard Country], "United States") then "US -Gulf"
        else if Text.Contains([Project Name], "NV-") and Text.Contains([Standard Country], "United States") then "US -West"
        else if Text.Contains([WBS Description], "PA-") and Text.Contains([Standard Country], "United States") then "US -North"
        else if Text.Contains([WBS Description], "NC-") and Text.Contains([Standard Country], "United States") then "US -South"
        else if Text.Contains([WBS Description], "TX-") and Text.Contains([Standard Country], "United States") then "US -Gulf"
        else if Text.Contains([WBS Description], "NV-") and Text.Contains([Standard Country], "United States") then "US -West"
        else [Standard Country]
    ),
    
    // Clean up country columns
    RemoveOldStandardCountry = Table.RemoveColumns(AddCountryEnhanced, {"Standard Country"}),
    RenameCountryEnhanced = Table.RenameColumns(RemoveOldStandardCountry, {
        {"Country_en", "Standard Country"}, 
        {"Total Amount USD", "Sales Amount"}
    }),
    
    // Format Month as two-digit string
    SetMonthYearTypes = Table.TransformColumnTypes(RenameCountryEnhanced, {
        {"Month", type text}, 
        {"Year", type text}
    }),
    AddMonthFormatted = Table.AddColumn(SetMonthYearTypes, "Month_N", each Text.PadStart(Text.From([Month]), 2, "0")),
    RemoveOldMonthCol = Table.RemoveColumns(AddMonthFormatted, {"Month"}),
    RenameMonthFormatted = Table.RenameColumns(RemoveOldMonthCol, {{"Month_N", "Month"}}),
    SetMonthTextType = Table.TransformColumnTypes(RenameMonthFormatted, {{"Month", type text}}),
    
    // Create YearMonthDate composite key (YYYYMMDD format, first day of month)
    AddYearMonthDateKey = Table.AddColumn(SetMonthTextType, "YearMonthDate", each [Year] & [Month] & "01"),
    SetYearMonthDateType = Table.TransformColumnTypes(AddYearMonthDateKey, {{"YearMonthDate", type text}}),
    
    // Add Parent/Child WBS flag
    // Business Rule: WBS Numbers containing "-C-" are Child WBS elements
    AddParentChildFlag = Table.AddColumn(SetYearMonthDateType, "ParentChildFlag", each 
        if Text.Contains([WBS Number], "-C-") then "Child" else "Parent"
    ),
    
    // Calculate Billbacks value (Child WBS only)
    AddBillbacksValue = Table.AddColumn(AddParentChildFlag, "Billbacks_Value", each 
        if [ParentChildFlag] = "Child" then [Sales Amount] else 0
    ),
    
    // Calculate enhanced Sales Amount (Parent WBS only)
    AddSalesAmountEnhanced = Table.AddColumn(AddBillbacksValue, "Sales Amount_enh", each 
        if [ParentChildFlag] = "Parent" then [Sales Amount] else 0
    ),
    
    // Reorder columns for logical grouping
    ReorderFinalColumns = Table.ReorderColumns(AddSalesAmountEnhanced, {
        "Project Number", "Project Name", "Region_Delete", "Country_Delete", 
        "Internal/External", "Sold To", "Sold to Description", "Responsible Person", 
        "Company Code", "Order Type", "Business", "Product Line", "PHx DS 17 PLs", 
        "WBS Number", "WBS Description", "WBS Product Line", "Month", "Quarter", "Year", 
        "Sales Amount", "Sales Amount_enh", "Standard Region", "Standard Sub Region", 
        "PH1", "PH2", "PH3", "PH4", "PH5", "Revenue Stream", "Delivery Mode", 
        "SMF Mapping 1", "SMF Mapping 2", "Order Type Category", "Data Source", 
        "INT/EXT Type", "Services PL Flag", "Standard Country", "YearMonthDate", 
        "ParentChildFlag", "Billbacks_Value"
    }),
    
    // Use enhanced Sales Amount as primary, remove original
    SetSalesAmountEnhType = Table.TransformColumnTypes(ReorderFinalColumns, {{"Sales Amount_enh", type number}}),
    RemoveOldSalesAmount = Table.RemoveColumns(SetSalesAmountEnhType, {"Sales Amount"}),
    RenameSalesAmountEnh = Table.RenameColumns(RemoveOldSalesAmount, {{"Sales Amount_enh", "Sales Amount"}}),
    SetBillbacksType = Table.TransformColumnTypes(RenameSalesAmountEnh, {{"Billbacks_Value", type number}}),
    
    // Add Quarter and Year copy columns for current period comparison
    DuplicateQuarter = Table.DuplicateColumn(SetBillbacksType, "Quarter", "Quarter - Copy"),
    SetQuarterCopyType = Table.TransformColumnTypes(DuplicateQuarter, {{"Quarter - Copy", type text}}),
    DuplicateYear = Table.DuplicateColumn(SetQuarterCopyType, "Year", "Year - Copy"),
    
    // Add current quarter/year for comparison
    AddCurrentQuarter = Table.AddColumn(DuplicateYear, "CurrentQuarter", each Text.From(Date.QuarterOfYear(DateTime.LocalNow()))),
    SetCurrentQuarterType = Table.TransformColumnTypes(AddCurrentQuarter, {{"CurrentQuarter", type text}}),
    AddCurrentYear = Table.AddColumn(SetCurrentQuarterType, "CurrentYear", each Text.From(Date.Year(DateTime.LocalNow()))),
    SetCurrentYearType = Table.TransformColumnTypes(AddCurrentYear, {{"CurrentYear", type text}}),
    
    // Flag current quarter records
    AddCurrentQuarterFlag = Table.AddColumn(SetCurrentYearType, "Current Quarter", each 
        if [#"Year - Copy"] = [CurrentYear] and [#"Quarter - Copy"] = [CurrentQuarter] then "Yes" else "No"
    ),
    
    // Join with Internal External Standard Map for Sales
    JoinIntExtMap = Table.NestedJoin(AddCurrentQuarterFlag, {"Internal/External"}, #"Internal External Standard Map_Sales", {"Sales Report"}, "Internal External Standard Map", JoinKind.LeftOuter),
    ExpandIntExtMap = Table.ExpandTableColumn(JoinIntExtMap, "Internal External Standard Map", {"Standard Internal/External "}, {"Internal External Standard Map.Standard Internal/External "}),
    
    // Join with End User mapping
    JoinEndUserMap = Table.NestedJoin(ExpandIntExtMap, {"WBS Number"}, EndUserMapTBL, {"WBS_EndUser"}, "EndUserMapTBL", JoinKind.LeftOuter),
    ExpandEndUserMap = Table.ExpandTableColumn(JoinEndUserMap, "EndUserMapTBL", {"End User"}, {"End User"}),
    
    // Handle PROJECT TYPE errors
    ReplaceEmptyProjectType = Table.ReplaceValue(ExpandEndUserMap, "", "Not Found", Replacer.ReplaceValue, {"PROJECT TYPE"}),
    ReplaceEmptyProjectType2 = Table.ReplaceValue(ReplaceEmptyProjectType, "", "Not Found", Replacer.ReplaceValue, {"PROJECT TYPE"}),
    
    // Process End User columns
    RenameEndUser1 = Table.RenameColumns(ReplaceEmptyProjectType2, {{"End User", "End User1"}}),
    ReplaceNullEndUser1 = Table.ReplaceValue(RenameEndUser1, null, "Not Found", Replacer.ReplaceValue, {"End User1"}),
    ReplaceNullEndUserEx = Table.ReplaceValue(ReplaceNullEndUser1, null, "Not Found", Replacer.ReplaceValue, {"End user_Ex"}),
    
    // Combine End User sources (prefer End User1 from mapping, fallback to End user_Ex)
    AddEndUserFinal = Table.AddColumn(ReplaceNullEndUserEx, "End User", each 
        if [End User1] = "Not Found" then [End user_Ex] 
        else if [End user_Ex] = "Not Found" then [End User1] 
        else [End User1]
    ),
    
    // Handle Standard Country errors
    HandleCountryErrors2 = Table.ReplaceErrorValues(AddEndUserFinal, {{"Standard Country", "Not assigned"}}),
    ReplaceNullStandardCountry = Table.ReplaceValue(HandleCountryErrors2, null, "Not assigned", Replacer.ReplaceValue, {"Standard Country"}),
    
    // Add WBS Level based on phase codes
    // Business Rule: WBS phases P-01 through P-08 and C-01 through C-08 indicate project phases
    AddWBSLevel = Table.AddColumn(ReplaceNullStandardCountry, "WBS Level", each 
        let
            Phases = {"P-01", "P-02", "P-03", "P-07", "P-08", "C-01", "C-02", "C-03", "C-07", "C-08"},
            Match = List.Select(Phases, (p) => Text.Contains([WBS Number], p, Comparer.OrdinalIgnoreCase))
        in
            if List.Count(Match) > 0 then Match{0} else "Not Assigned"
    ),
    SetWBSLevelType = Table.TransformColumnTypes(AddWBSLevel, {{"WBS Level", type text}}),
    HandleWBSLevelErrors = Table.ReplaceErrorValues(SetWBSLevelType, {{"WBS Level", "error"}})
in
    HandleWBSLevelErrors


// ============================================================================
// QUERY: Internal External Standard Map_Order
// PURPOSE: Internal/External classification mapping for Orders reports
// SOURCE: SMF Mapping File Master - Internal External Standard Map sheet
// BUSINESS USE: Standardizes customer type classification for Orders reporting
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    MappingSheet = Source{[Item="Internal External Standard Map",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(MappingSheet, [PromoteAllScalars=true]),
    
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
    
    // Filter valid rows
    FilterValidStandard = Table.SelectRows(SetColumnTypes, each [#"Standard Internal/External "] <> null and [#"Standard Internal/External "] <> ""),
    
    // Remove columns not needed for Order report mapping
    RemoveOtherReportColumns = Table.RemoveColumns(FilterValidStandard, {
        "Sales Report", "Backlog Report", "Convertible Order", 
        "As Sold Margin", "As Sold Report", "PLM", "Unbilled BDR", "Sales WORP"
    }),
    
    // Filter to rows with Order Report mapping
    FilterValidOrderReport = Table.SelectRows(RemoveOtherReportColumns, each [Order Report] <> null and [Order Report] <> "")
in
    FilterValidOrderReport


// ============================================================================
// QUERY: Internal External Standard Map_Backlog
// PURPOSE: Internal/External classification mapping for Backlog reports
// SOURCE: SMF Mapping File Master - Internal External Standard Map sheet
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    MappingSheet = Source{[Item="Internal External Standard Map",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(MappingSheet, [PromoteAllScalars=true]),
    
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
    
    FilterValidStandard = Table.SelectRows(SetColumnTypes, each [#"Standard Internal/External "] <> null and [#"Standard Internal/External "] <> ""),
    RemoveOtherReportColumns = Table.RemoveColumns(FilterValidStandard, {
        "Sales Report", "Order Report", "Convertible Order", 
        "As Sold Margin", "As Sold Report", "PLM", "Unbilled BDR", "Sales WORP"
    }),
    FilterValidBacklogReport = Table.SelectRows(RemoveOtherReportColumns, each [Backlog Report] <> null and [Backlog Report] <> "")
in
    FilterValidBacklogReport


// ============================================================================
// QUERY: Internal External Standard Map_CO
// PURPOSE: Internal/External classification mapping for Convertible Orders
// SOURCE: SMF Mapping File Master - Internal External Standard Map sheet
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    MappingSheet = Source{[Item="Internal External Standard Map",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(MappingSheet, [PromoteAllScalars=true]),
    
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
    
    FilterValidStandard = Table.SelectRows(SetColumnTypes, each [#"Standard Internal/External "] <> null and [#"Standard Internal/External "] <> ""),
    RemoveOtherReportColumns = Table.RemoveColumns(FilterValidStandard, {
        "Sales Report", "Order Report", "Backlog Report", 
        "As Sold Margin", "As Sold Report", "PLM", "Unbilled BDR", "Sales WORP"
    }),
    FilterValidCOReport = Table.SelectRows(RemoveOtherReportColumns, each [Convertible Order] <> null and [Convertible Order] <> "")
in
    FilterValidCOReport


// ============================================================================
// QUERY: Internal External Standard Map_Assold
// PURPOSE: Internal/External classification mapping for As Sold reports
// SOURCE: SMF Mapping File Master - Internal External Standard Map sheet
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    MappingSheet = Source{[Item="Internal External Standard Map",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(MappingSheet, [PromoteAllScalars=true]),
    
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
    
    FilterValidStandard = Table.SelectRows(SetColumnTypes, each [#"Standard Internal/External "] <> null and [#"Standard Internal/External "] <> ""),
    RemoveOtherReportColumns = Table.RemoveColumns(FilterValidStandard, {
        "Sales Report", "Order Report", "Backlog Report", 
        "Convertible Order", "PLM", "Unbilled BDR", "Sales WORP"
    }),
    FilterValidAsSoldReport = Table.SelectRows(RemoveOtherReportColumns, each [#"As Sold Report"] <> null and [#"As Sold Report"] <> "")
in
    FilterValidAsSoldReport


// ============================================================================
// QUERY: Internal External Standard Map_Worp
// PURPOSE: Internal/External classification mapping for WORP reports
// SOURCE: SMF Mapping File Master - Internal External Standard Map sheet
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Mapping%20Files%20-Do%20Not%20Delete/SMF%20Mapping%20File-Master.xlsx"), 
        null, 
        true
    ),
    
    MappingSheet = Source{[Item="Internal External Standard Map",Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(MappingSheet, [PromoteAllScalars=true]),
    
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
    
    FilterValidStandard = Table.SelectRows(SetColumnTypes, each [#"Standard Internal/External "] <> null and [#"Standard Internal/External "] <> ""),
    RemoveOtherReportColumns = Table.RemoveColumns(FilterValidStandard, {
        "Sales Report", "Order Report", "Backlog Report", 
        "Convertible Order", "As Sold Margin", "As Sold Report", "PLM", "Unbilled BDR"
    }),
    FilterValidWorpReport = Table.SelectRows(RemoveOtherReportColumns, each [Sales WORP] <> null and [Sales WORP] <> "")
in
    FilterValidWorpReport


// ============================================================================
// QUERY: AR Data
// PURPOSE: Load Accounts Receivable and Past Due data
// SOURCE: BN AR_Unbilled_Open BMS master file
// BUSINESS USE: AR aging analysis and past due tracking
// ============================================================================
let
    // Connect to AR data source
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/SMF%20Unbilled%20PD%20AR/Current%20Week/BN%20AR_Unbilled_Open%20BMS%20master%20file.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to AR Data sheet
    ARDataSheet = Source{[Item="AR Data",Kind="Sheet"]}[Data],
    
    // Remove extra/unused columns (system columns beyond the data)
    RemoveExtraColumns = Table.RemoveColumns(ARDataSheet, {
        "Column157", "Column158", "Column159", "Column160", "Column161", 
        "Column162", "Column163", "Column164", "Column165", "Column166", 
        "Column167", "Column168", "Column169", "Column170", "Column171", 
        "Column172", "Column173", "Column174", "Column29"
    }),
    
    // Promote headers
    PromotedHeaders = Table.PromoteHeaders(RemoveExtraColumns, [PromoteAllScalars=true]),
    
    // Standardize Customer Type values
    // Business Rule: Map EXTERNAL to External, RELATED PARTY to Internal
    ReplaceExternalType = Table.ReplaceValue(PromotedHeaders, "EXTERNAL", "External", Replacer.ReplaceText, {"Customer Type"}),
    ReplaceInternalType = Table.ReplaceValue(ReplaceExternalType, "RELATED PARTY", "Internal", Replacer.ReplaceText, {"Customer Type"}),
    
    // Set numeric types for financial columns
    SetNumericTypes = Table.TransformColumnTypes(ReplaceInternalType, {
        {"PD", type number}, 
        {"AR", type number}
    }),
    
    // Rename Sales Office Region for clarity
    RenameSalesOffice = Table.RenameColumns(SetNumericTypes, {{"Sales Office Region", "DS SO region"}}),
    
    // Filter to relevant service types (OCSR1/OCSR2), exclude Repair and GE AERO
    // Business Rule: Only include core service records, not repairs or aerospace
    FilterServiceTypes = Table.SelectRows(RenameSalesOffice, each 
        ([#"PRJ/SRV"] = "OCSR1" or [#"PRJ/SRV"] = "OCSR2") and 
        ([Repair] = "No") and 
        ([GE AERO] = "No")
    ),
    
    // Exclude ARMS records
    FilterExcludeARMS = Table.SelectRows(FilterServiceTypes, each [ARMS] = "No")
in
    FilterExcludeARMS



// ============================================================================
// QUERY: WORP_2020Orders
// PURPOSE: Load 2020 WORP (Worldwide Orders Reporting Platform) data
// SOURCE: Digital Solutions YTD Transaction File
// BUSINESS USE: Historical WORP orders for trend analysis
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/WORP/Digital%20Solutions%20YTD_Transaction%20File%20_Week%2052%20Final%20Revised%202020%20-%20Channel%20info.xlsx"), 
        null, 
        true
    ),
    
    // Navigate to 2020 Orders sheet
    OrdersSheet = Source{[Name="2020 Orders"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(OrdersSheet, [PromoteAllScalars=true]),
    
    // Set all column types
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"CRV-Marvel-SOURCE", type text}, 
        {"SOR Name", type text}, 
        {"Unique Identifier", type text}, 
        {"Order Number", type text}, 
        {"Fiscal Wk", Int64.Type}, 
        {"Fiscal Qtr", Int64.Type}, 
        {"Fiscal Yr", Int64.Type}, 
        {"Transaction Type", type text}, 
        {"P&L Tier1", type text}, 
        {"P&L Tier2", type text}, 
        {"P&L TIER 3", type text}, 
        {"P&L TIER 4", type text}, 
        {"P&L TIER 5", type text}, 
        {"SALES SEGMENT NAME", type text}, 
        {"P&L-CRV INTERNAL", type text}, 
        {"Order/Commitment Amount", type number}, 
        {"GGO region name", type text}, 
        {"P&L region name", type text}, 
        {"Sold to Country", type text}, 
        {"Sold to State/Location", type text}, 
        {"Site Country Name", type text}, 
        {"Order Date", type date}, 
        {"Order Description", type text}, 
        {"Product Type Description", type text}, 
        {"Offering Family Name", type text}, 
        {"Project Name", type text}, 
        {"Project Type", type text}
    }),
    
    // Select only relevant columns
    SelectColumns = Table.SelectColumns(SetColumnTypes, {
        "CRV-Marvel-SOURCE", "SOR Name", "Unique Identifier", "Order Number", 
        "Transaction Type", "P&L Tier1", "P&L Tier2", "P&L TIER 3", "P&L TIER 4", 
        "P&L TIER 5", "SALES SEGMENT NAME", "P&L-CRV INTERNAL", "Order/Commitment Amount", 
        "GGO region name", "P&L region name", "Sold to Country", "Sold to State/Location", 
        "Site Country Name", "Order Date", "Order Description", "Product Type Description", 
        "Offering Family Name", "Project Name", "Project Type"
    }),
    
    // Filter to BN and Manual OC sources only
    FilterSORName = Table.SelectRows(SelectColumns, each 
        ([SOR Name] = "BW-BN" or [SOR Name] = "ManualOC")
    ),
    
    // Uppercase tier columns for consistency
    UppercaseTiers = Table.TransformColumns(FilterSORName, {
        {"P&L TIER 3", Text.Upper, type text}, 
        {"P&L Tier2", Text.Upper, type text}
    }),
    
    // Filter to Project type only (P)
    FilterProjectType = Table.SelectRows(UppercaseTiers, each ([Project Type] = "P"))
in
    FilterProjectType


// ============================================================================
// QUERY: WORP_2021Orders
// PURPOSE: Load 2021 WORP orders data
// SOURCE: 2021 WORP Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/WORP/2021%20WORP.xlsx"), 
        null, 
        true
    ),
    
    OrdersSheet = Source{[Name="2021 Total BN"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(OrdersSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"CRV-Marvel-SOURCE", type text}, 
        {"SOR Name", type text}, 
        {"Unique Identifier", type text}, 
        {"Order Number", type text}, 
        {"Transaction Type", type text}, 
        {"P&L Tier1", type text}, 
        {"P&L Tier2", type text}, 
        {"P&L TIER 3", type text}, 
        {"P&L TIER 4", type text}, 
        {"P&L TIER 5", type text}, 
        {"SALES SEGMENT NAME", type text}, 
        {"P&L-CRV INTERNAL", type text}, 
        {"Order/Commitment Amount", type number}, 
        {"GGO region name", type text}, 
        {"P&L region name", type text}, 
        {"Sold to Country", type text}, 
        {"Sold to State/Location", type text}, 
        {"Site Country Name", type text}, 
        {"Order Date", type date}, 
        {"Order Description", type text}, 
        {"Product Type Description", type text}, 
        {"Offering Family Name", type text}, 
        {"Project Name", type text}, 
        {"Project Type", type text}
    }),
    
    SelectColumns = Table.SelectColumns(SetColumnTypes, {
        "CRV-Marvel-SOURCE", "SOR Name", "Unique Identifier", "Order Number", 
        "Transaction Type", "P&L Tier1", "P&L Tier2", "P&L TIER 3", "P&L TIER 4", 
        "P&L TIER 5", "SALES SEGMENT NAME", "P&L-CRV INTERNAL", "Order/Commitment Amount", 
        "GGO region name", "P&L region name", "Sold to Country", "Sold to State/Location", 
        "Site Country Name", "Order Date", "Order Description", "Product Type Description", 
        "Offering Family Name", "Project Name", "Project Type"
    }),
    
    FilterSORName = Table.SelectRows(SelectColumns, each 
        ([SOR Name] = "BW-BN" or [SOR Name] = "ManualOC")
    ),
    
    UppercaseTiers = Table.TransformColumns(FilterSORName, {
        {"P&L TIER 3", Text.Upper, type text}, 
        {"P&L Tier2", Text.Upper, type text}
    }),
    
    FilterProjectType = Table.SelectRows(UppercaseTiers, each ([Project Type] = "P"))
in
    FilterProjectType


// ============================================================================
// QUERY: WORP_2022Orders
// PURPOSE: Load 2022 WORP orders data
// SOURCE: 2022 WORP Data Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/WORP/2022%20WORP%20Data.xlsx"), 
        null, 
        true
    ),
    
    OrdersSheet = Source{[Name="2022 Total orders"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(OrdersSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"CRV-Marvel-SOURCE", type text}, 
        {"SOR Name", type text}, 
        {"Unique Identifier", type text}, 
        {"Order Number", type text}, 
        {"Transaction Type", type text}, 
        {"P&L Tier1", type text}, 
        {"P&L Tier2", type text}, 
        {"P&L TIER 3", type text}, 
        {"P&L TIER 4", type text}, 
        {"P&L TIER 5", type text}, 
        {"SALES SEGMENT NAME", type text}, 
        {"P&L-CRV INTERNAL", type text}, 
        {"Order/Commitment Amount", type number}, 
        {"GGO region name", type text}, 
        {"P&L region name", type text}, 
        {"Sold to Country", type text}, 
        {"Sold to State/Location", type text}, 
        {"Site Country Name", type text}, 
        {"Order Date", type date}, 
        {"Order Description", type text}, 
        {"Product Type Description", type text}, 
        {"Offering Family Name", type text}, 
        {"Project Name", type text}, 
        {"Project Type", type text}
    }),
    
    SelectColumns = Table.SelectColumns(SetColumnTypes, {
        "CRV-Marvel-SOURCE", "SOR Name", "Unique Identifier", "Order Number", 
        "Transaction Type", "P&L Tier1", "P&L Tier2", "P&L TIER 3", "P&L TIER 4", 
        "P&L TIER 5", "SALES SEGMENT NAME", "P&L-CRV INTERNAL", "Order/Commitment Amount", 
        "GGO region name", "P&L region name", "Sold to Country", "Sold to State/Location", 
        "Site Country Name", "Order Date", "Order Description", "Product Type Description", 
        "Offering Family Name", "Project Name", "Project Type"
    }),
    
    FilterSORName = Table.SelectRows(SelectColumns, each 
        ([SOR Name] = "BW-BN" or [SOR Name] = "ManualOC")
    ),
    
    UppercaseTiers = Table.TransformColumns(FilterSORName, {
        {"P&L TIER 3", Text.Upper, type text}, 
        {"P&L Tier2", Text.Upper, type text}
    }),
    
    FilterProjectType = Table.SelectRows(UppercaseTiers, each ([Project Type] = "P"))
in
    FilterProjectType


// ============================================================================
// QUERY: WORP_2023Orders
// PURPOSE: Load 2023 WORP orders data
// SOURCE: 2023 WORP Data Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/WORP/2023%20WORP%20Data.xlsx"), 
        null, 
        true
    ),
    
    OrdersSheet = Source{[Name="2023 Total orders"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(OrdersSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"CRV-Marvel-SOURCE", type text}, 
        {"SOR Name", type text}, 
        {"Unique Identifier", type text}, 
        {"Order Number", type text}, 
        {"Transaction Type", type text}, 
        {"P&L Tier1", type text}, 
        {"P&L Tier2", type text}, 
        {"P&L TIER 3", type text}, 
        {"P&L TIER 4", type text}, 
        {"P&L TIER 5", type text}, 
        {"SALES SEGMENT NAME", type text}, 
        {"P&L-CRV INTERNAL", type text}, 
        {"Order/Commitment Amount", type number}, 
        {"GGO region name", type text}, 
        {"P&L region name", type text}, 
        {"Sold to Country", type text}, 
        {"Sold to State/Location", type text}, 
        {"Site Country Name", type text}, 
        {"Order Date", type date}, 
        {"Order Description", type text}, 
        {"Product Type Description", type text}, 
        {"Offering Family Name", type text}, 
        {"Project Name", type text}, 
        {"Project Type", type text}
    }),
    
    SelectColumns = Table.SelectColumns(SetColumnTypes, {
        "CRV-Marvel-SOURCE", "SOR Name", "Unique Identifier", "Order Number", 
        "Transaction Type", "P&L Tier1", "P&L Tier2", "P&L TIER 3", "P&L TIER 4", 
        "P&L TIER 5", "SALES SEGMENT NAME", "P&L-CRV INTERNAL", "Order/Commitment Amount", 
        "GGO region name", "P&L region name", "Sold to Country", "Sold to State/Location", 
        "Site Country Name", "Order Date", "Order Description", "Product Type Description", 
        "Offering Family Name", "Project Name", "Project Type"
    }),
    
    FilterSORName = Table.SelectRows(SelectColumns, each 
        ([SOR Name] = "BW-BN" or [SOR Name] = "ManualOC")
    ),
    
    UppercaseTiers = Table.TransformColumns(FilterSORName, {
        {"P&L TIER 3", Text.Upper, type text}, 
        {"P&L Tier2", Text.Upper, type text}
    }),
    
    FilterProjectType = Table.SelectRows(UppercaseTiers, each ([Project Type] = "P"))
in
    FilterProjectType


// ============================================================================
// QUERY: WORP_2024Orders
// PURPOSE: Load 2024 WORP orders data
// SOURCE: 2024 WORP Data Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/WORP/2024%20WORP%20Data.xlsx"), 
        null, 
        true
    ),
    
    OrdersSheet = Source{[Name="2024 Total orders"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(OrdersSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"CRV-Marvel-SOURCE", type text}, 
        {"SOR Name", type text}, 
        {"Unique Identifier", type text}, 
        {"Order Number", type text}, 
        {"Transaction Type", type text}, 
        {"P&L Tier1", type text}, 
        {"P&L Tier2", type text}, 
        {"P&L TIER 3", type text}, 
        {"P&L TIER 4", type text}, 
        {"P&L TIER 5", type text}, 
        {"SALES SEGMENT NAME", type text}, 
        {"P&L-CRV INTERNAL", type text}, 
        {"Order/Commitment Amount", type number}, 
        {"GGO region name", type text}, 
        {"P&L region name", type text}, 
        {"Sold to Country", type text}, 
        {"Sold to State/Location", type text}, 
        {"Site Country Name", type text}, 
        {"Order Date", type date}, 
        {"Order Description", type text}, 
        {"Product Type Description", type text}, 
        {"Offering Family Name", type text}, 
        {"Project Name", type text}, 
        {"Project Type", type text}
    }),
    
    SelectColumns = Table.SelectColumns(SetColumnTypes, {
        "CRV-Marvel-SOURCE", "SOR Name", "Unique Identifier", "Order Number", 
        "Transaction Type", "P&L Tier1", "P&L Tier2", "P&L TIER 3", "P&L TIER 4", 
        "P&L TIER 5", "SALES SEGMENT NAME", "P&L-CRV INTERNAL", "Order/Commitment Amount", 
        "GGO region name", "P&L region name", "Sold to Country", "Sold to State/Location", 
        "Site Country Name", "Order Date", "Order Description", "Product Type Description", 
        "Offering Family Name", "Project Name", "Project Type"
    }),
    
    FilterSORName = Table.SelectRows(SelectColumns, each 
        ([SOR Name] = "BW-BN" or [SOR Name] = "ManualOC")
    ),
    
    UppercaseTiers = Table.TransformColumns(FilterSORName, {
        {"P&L TIER 3", Text.Upper, type text}, 
        {"P&L Tier2", Text.Upper, type text}
    }),
    
    FilterProjectType = Table.SelectRows(UppercaseTiers, each ([Project Type] = "P"))
in
    FilterProjectType


// ============================================================================
// QUERY: WORP_2025Orders
// PURPOSE: Load 2025 WORP orders data (current year)
// SOURCE: 2025 WORP Data Excel file
// ============================================================================
let
    Source = Excel.Workbook(
        Web.Contents("https://bakerhughes.sharepoint.com/sites/IETCordantOperationsDashboard/Shared%20Documents/Dashboard%20Data%20Sources/FINEX%20Data%20Sources/FINEX_ALL_SOURCES/WORP/2025%20WORP%20Data.xlsx"), 
        null, 
        true
    ),
    
    OrdersSheet = Source{[Name="2025 Total orders"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(OrdersSheet, [PromoteAllScalars=true]),
    
    SetColumnTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"CRV-Marvel-SOURCE", type text}, 
        {"SOR Name", type text}, 
        {"Unique Identifier", type text}, 
        {"Order Number", type text}, 
        {"Transaction Type", type text}, 
        {"P&L Tier1", type text}, 
        {"P&L Tier2", type text}, 
        {"P&L TIER 3", type text}, 
        {"P&L TIER 4", type text}, 
        {"P&L TIER 5", type text}, 
        {"SALES SEGMENT NAME", type text}, 
        {"P&L-CRV INTERNAL", type text}, 
        {"Order/Commitment Amount", type number}, 
        {"GGO region name", type text}, 
        {"P&L region name", type text}, 
        {"Sold to Country", type text}, 
        {"Sold to State/Location", type text}, 
        {"Site Country Name", type text}, 
        {"Order Date", type date}, 
        {"Order Description", type text}, 
        {"Product Type Description", type text}, 
        {"Offering Family Name", type text}, 
        {"Project Name", type text}, 
        {"Project Type", type text}
    }),
    
    SelectColumns = Table.SelectColumns(SetColumnTypes, {
        "CRV-Marvel-SOURCE", "SOR Name", "Unique Identifier", "Order Number", 
        "Transaction Type", "P&L Tier1", "P&L Tier2", "P&L TIER 3", "P&L TIER 4", 
        "P&L TIER 5", "SALES SEGMENT NAME", "P&L-CRV INTERNAL", "Order/Commitment Amount", 
        "GGO region name", "P&L region name", "Sold to Country", "Sold to State/Location", 
        "Site Country Name", "Order Date", "Order Description", "Product Type Description", 
        "Offering Family Name", "Project Name", "Project Type"
    }),
    
    FilterSORName = Table.SelectRows(SelectColumns, each 
        ([SOR Name] = "BW-BN" or [SOR Name] = "ManualOC")
    ),
    
    UppercaseTiers = Table.TransformColumns(FilterSORName, {
        {"P&L TIER 3", Text.Upper, type text}, 
        {"P&L Tier2", Text.Upper, type text}
    }),
    
    FilterProjectType = Table.SelectRows(UppercaseTiers, each ([Project Type] = "P"))
in
    FilterProjectType


// ============================================================================
// QUERY: WORP_Orders
// PURPOSE: Combined WORP orders from all years with standardization
// SOURCE: Appended from WORP_2020Orders through WORP_2025Orders
// BUSINESS USE: Complete WORP orders dataset for multi-year analysis
// DEPENDENCIES: WORP_2020Orders through WORP_2025Orders, Standard Country List-WORP,
//               DM_StandardMapping, Internal External Standard Map_Worp
// ============================================================================
let
    // Combine all WORP year tables
    Source = Table.Combine({
        WORP_2022Orders, 
        WORP_2020Orders, 
        WORP_2021Orders, 
        WORP_2023Orders, 
        WORP_2024Orders, 
        WORP_2025Orders
    }),
    
    // Rename columns to standard names
    RenameOrderAmount = Table.RenameColumns(Source, {{"Order/Commitment Amount", "Orders Amount"}}),
    RenameRegion = Table.RenameColumns(RenameOrderAmount, {{"P&L region name", "Region"}}),
    RenameCountry = Table.RenameColumns(RenameRegion, {{"Site Country Name", "Country"}}),
    RenameInternal = Table.RenameColumns(RenameCountry, {{"P&L-CRV INTERNAL", "Internal/External"}}),
    
    // Join with Standard Country List for geographic standardization
    JoinCountryList = Table.NestedJoin(RenameInternal, {"Region", "Country"}, #"Standard Country List-WORP", {"Region-WORP", "Country-WORP"}, "Standard Country List-WORP", JoinKind.LeftOuter),
    ExpandCountryList = Table.ExpandTableColumn(JoinCountryList, "Standard Country List-WORP", {"Standard Region", "Standard Sub Region", "Standard Country"}, {"Standard Region", "Standard Sub Region", "Standard Country"}),
    
    // Join with DM Standard Mapping for product standardization
    JoinDMMapping = Table.NestedJoin(ExpandCountryList, {"P&L TIER 3"}, DM_StandardMapping, {"DM PL Mapping"}, "DM_StandardMapping", JoinKind.LeftOuter),
    ExpandDMMapping = Table.ExpandTableColumn(JoinDMMapping, "DM_StandardMapping", {"Standard PL"}, {"Standard PL"}),
    
    // Handle null Standard PL
    ReplaceNullStandardPL = Table.ReplaceValue(ExpandDMMapping, null, "Not found", Replacer.ReplaceValue, {"Standard PL"}),
    
    // Join with Internal/External mapping for WORP
    JoinIntExtMapping = Table.NestedJoin(ReplaceNullStandardPL, {"Internal/External"}, #"Internal External Standard Map_Worp", {"Sales WORP"}, "Internal External Standard Map", JoinKind.LeftOuter),
    ExpandIntExtMapping = Table.ExpandTableColumn(JoinIntExtMapping, "Internal External Standard Map", {"Standard Internal/External "}, {"Standard Internal/External"})
in
    ExpandIntExtMapping

