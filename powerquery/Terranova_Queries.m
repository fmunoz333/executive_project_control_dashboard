// =============================================================================
// Terranova 4G 100 - Power Query (M) queries
// Workbook : Terranova_100_4G_Excel_Dashboard.xlsx
// =============================================================================

// ===== QUERY: pStatusDate  (parameter read from Control!C6) ==================
let
    Source = Excel.CurrentWorkbook(){[Name = "StatusDate"]}[Content],
    Value  = Date.From(Source{0}[Column1])
in
    Value


// ===== QUERY: pGreenThr  (SPI threshold for GREEN, Control!C9) ===============
let
    Source = Excel.CurrentWorkbook(){[Name = "GreenThr"]}[Content],
    Value  = Number.From(Source{0}[Column1])
in
    Value


// ===== QUERY: pAmberThr  (SPI threshold for AMBER, Control!C10) ==============
let
    Source = Excel.CurrentWorkbook(){[Name = "AmberThr"]}[Content],
    Value  = Number.From(Source{0}[Column1])
in
    Value


// ===== QUERY: fct_Milestones  (clean, typed fact table) ======================
// Why: formulas that return "" arrive as text. Converting them to null and
// typing every column lets the Data Model treat FinishReal as a true date.
let
    Source      = Excel.CurrentWorkbook(){[Name = "FactMilestones"]}[Content],
    KeepColumns = Table.SelectColumns(Source, {
        "DateID", "SiteID", "SiteCode", "Province", "Municipality", "Owner",
        "AccessType", "StructureType", "Milestone", "MilestoneNo", "FinishPlan",
        "FinishReal", "SlipDays", "Status", "SiteBudgetUSD", "MilestoneWeightUSD",
        "CompletedFlag", "LateFlag", "SlipDaysNum"}),
    BlankToNull = Table.ReplaceValue(KeepColumns, "", null, Replacer.ReplaceValue,
                      {"FinishReal", "SlipDays"}),
    Typed       = Table.TransformColumnTypes(BlankToNull, {
        {"DateID", Int64.Type}, {"SiteID", Int64.Type}, {"SiteCode", type text},
        {"Province", type text}, {"Municipality", type text}, {"Owner", type text},
        {"AccessType", type text}, {"StructureType", type text}, {"Milestone", type text},
        {"MilestoneNo", Int64.Type}, {"FinishPlan", type date}, {"FinishReal", type date},
        {"SlipDays", Int64.Type}, {"Status", type text}, {"SiteBudgetUSD", type number},
        {"MilestoneWeightUSD", type number}, {"CompletedFlag", Int64.Type},
        {"LateFlag", Int64.Type}, {"SlipDaysNum", Int64.Type}})
in
    Typed


// ===== QUERY: dim_Site  (one row per site) ====================================
let
    Source   = fct_Milestones,
    Columns  = Table.SelectColumns(Source, {"SiteID", "SiteCode", "Province",
                   "Municipality", "Owner", "AccessType", "StructureType", "SiteBudgetUSD"}),
    Distinct = Table.Distinct(Columns),
    Wave     = Table.AddColumn(Distinct, "RolloutWave", each Text.Start([SiteCode], 2), type text),
    Sorted   = Table.Sort(Wave, {{"SiteID", Order.Ascending}})
in
    Sorted


// ===== QUERY: dim_Milestone  (one row per milestone definition) ===============
let
    Source   = fct_Milestones,
    Columns  = Table.SelectColumns(Source, {"MilestoneNo", "Milestone"}),
    Distinct = Table.Distinct(Columns),
    Sorted   = Table.Sort(Distinct, {{"MilestoneNo", Order.Ascending}})
in
    Sorted


// ===== QUERY: dim_Calendar  (contiguous date table) ==========================
let
    StartDate = Date.StartOfMonth(List.Min(fct_Milestones[FinishPlan])),
    EndDate   = Date.EndOfMonth(List.Max(
                    List.Combine({fct_Milestones[FinishPlan], fct_Milestones[FinishReal]}))),
    DayCount  = Duration.Days(EndDate - StartDate) + 1,
    Dates     = List.Dates(StartDate, DayCount, #duration(1, 0, 0, 0)),
    AsTable   = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}, null, ExtraValues.Error),
    Typed     = Table.TransformColumnTypes(AsTable, {{"Date", type date}}),
    AddYear   = Table.AddColumn(Typed, "Year", each Date.Year([Date]), Int64.Type),
    AddQtr    = Table.AddColumn(AddYear, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
    AddMonNo  = Table.AddColumn(AddQtr, "MonthNo", each Date.Month([Date]), Int64.Type),
    AddMonNm  = Table.AddColumn(AddMonNo, "MonthName", each Date.ToText([Date], "MMM"), type text),
    AddMonSt  = Table.AddColumn(AddMonNm, "MonthStart", each Date.StartOfMonth([Date]), type date),
    AddYM     = Table.AddColumn(AddMonSt, "YearMonth", each Date.ToText([Date], "yyyy-MM"), type text)
in
    AddYM


// ===== QUERY: param_Control  (single-row parameter table for DAX) ============
let
    Source = #table(
        type table [StatusDate = date, GreenThr = number, AmberThr = number],
        {{pStatusDate, pGreenThr, pAmberThr}})
in
    Source


// ===== QUERY: Site_Performance_Scorecard  (NEW TABLE - one row per site) =====
// Earned-value scorecard at the status date. Sorted worst-first so the top rows
// are the recovery watch-list. Reconciles to the dashboard:
//   Sum(PV_USD) = EVM_SCurve!I8   Sum(EV_USD) = EVM_SCurve!I9
let
    Source  = fct_Milestones,
    Grouped = Table.Group(Source,
        {"SiteID", "SiteCode", "Province", "Municipality", "Owner", "AccessType"},
        {
            {"Milestones",    each Table.RowCount(_),             Int64.Type},
            {"Completed",     each List.Sum([CompletedFlag]),     Int64.Type},
            {"LateCount",     each List.Sum([LateFlag]),          Int64.Type},
            {"TotalSlipDays", each List.Sum([SlipDaysNum]),       Int64.Type},
            {"MaxSlipDays",   each List.Max([SlipDaysNum]),       Int64.Type},
            {"BAC_USD",       each List.Sum([MilestoneWeightUSD]), type number},
            {"PV_USD",        each List.Sum(Table.SelectRows(_,
                                  each [FinishPlan] <= pStatusDate)[MilestoneWeightUSD]), type number},
            {"EV_USD",        each List.Sum(Table.SelectRows(_,
                                  each [FinishReal] <> null and [FinishReal] <= pStatusDate)[MilestoneWeightUSD]), type number}
        }),
    NoNulls  = Table.ReplaceValue(Grouped, null, 0, Replacer.ReplaceValue, {"PV_USD", "EV_USD"}),
    AddComp  = Table.AddColumn(NoNulls, "CompletionPct",
                   each [Completed] / [Milestones], Percentage.Type),
    AddLate  = Table.AddColumn(AddComp, "LatePctOfCompleted",
                   each if [Completed] = 0 then null else [LateCount] / [Completed], Percentage.Type),
    AddSlip  = Table.AddColumn(AddLate, "AvgSlipDays",
                   each if [Completed] = 0 then null else [TotalSlipDays] / [Completed], type number),
    AddSV    = Table.AddColumn(AddSlip, "SV_USD", each [EV_USD] - [PV_USD], type number),
    AddSPI   = Table.AddColumn(AddSV, "SPI",
                   each if [PV_USD] = 0 then null else [EV_USD] / [PV_USD], type number),
    AddRAG   = Table.AddColumn(AddSPI, "RAG",
                   each if [SPI] = null then "N/A"
                        else if [SPI] >= pGreenThr then "GREEN"
                        else if [SPI] >= pAmberThr then "AMBER"
                        else "RED", type text),
    AddWave  = Table.AddColumn(AddRAG, "RolloutWave", each Text.Start([SiteCode], 2), type text),
    ByCost   = Table.Sort(AddWave, {{"BAC_USD", Order.Descending}}),
    Ranked   = Table.AddIndexColumn(ByCost, "InvestmentRank", 1, 1, Int64.Type),
    Final    = Table.Sort(Ranked, {{"SPI", Order.Ascending}, {"SV_USD", Order.Ascending}})
in
    Final
