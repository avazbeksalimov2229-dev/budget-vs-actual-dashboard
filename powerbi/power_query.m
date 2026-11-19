// Power Query (M) - Project Budget vs Actual Dashboard
// ------------------------------------------------------------------------------
// Parameter DataFolder (Text) = full path of the data folder ending with "\".
// Every type change uses the "en-US" culture (the CSV decimal separator is ".").
// Money columns use Currency.Type (fixed decimal), so totals add up to the cent.
//
// Sources
//   CSV exports   budget (estimating system), change orders, commitments, ERP job cost, billing
//   Excel files   one workbook per project manager in DataFolder\pm_updates\ (weekly % complete and
//                 monthly forecast to complete). New PM = new file in the folder, no query change.
// ------------------------------------------------------------------------------


// ---------- pm_workbooks ----------
// Staging query (not loaded): every PM workbook in the folder, except the blank template and Excel lock files.
let
    Files = Folder.Files(DataFolder & "pm_updates"),
    Workbooks = Table.SelectRows(Files, each [Extension] = ".xlsx"
        and not Text.StartsWith([Name], "~$")
        and not Text.StartsWith([Name], "PM_Update_TEMPLATE")),
    Opened = Table.AddColumn(Workbooks, "Workbook", each Excel.Workbook([Content], true)),
    Kept = Table.SelectColumns(Opened, {"Name", "Workbook"})
in
    Kept


// ---------- dim_project ----------
let
    Source = Csv.Document(File.Contents(DataFolder & "projects.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Headers, {
        {"project_id", type text}, {"project_name", type text}, {"client", type text}, {"project_type", type text},
        {"pm_id", type text}, {"start_date", type date}, {"planned_finish", type date},
        {"original_contract_value", Currency.Type}, {"status", type text}}, "en-US"),
    Label = Table.AddColumn(Typed, "project_label", each [project_id] & " " & [project_name], type text)
in
    Label


// ---------- dim_pm ----------
let
    Source = Csv.Document(File.Contents(DataFolder & "project_managers.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Headers, {{"pm_id", type text}, {"pm_name", type text}, {"pm_email", type text}})
in
    Typed


// ---------- dim_cost_code ----------
let
    Source = Csv.Document(File.Contents(DataFolder & "cost_codes.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Headers, {
        {"cost_code", type text}, {"cost_code_name", type text}, {"cost_category", type text},
        {"sort_order", Int64.Type}}, "en-US"),
    Label = Table.AddColumn(Typed, "cost_code_label", each [cost_code] & " " & [cost_code_name], type text)
in
    Label


// ---------- fact_budget ----------
// Monthly phased budget from the estimating system. Change order rows start in the month the CO was approved.
let
    Source = Csv.Document(File.Contents(DataFolder & "budget_phased.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    EmptyToNull = Table.ReplaceValue(Headers, "", null, Replacer.ReplaceValue, {"co_id"}),
    Typed = Table.TransformColumnTypes(EmptyToNull, {
        {"project_id", type text}, {"cost_code", type text}, {"month_start", type date},
        {"budget_type", type text}, {"co_id", type text}, {"amount", Currency.Type}}, "en-US")
in
    Typed


// ---------- fact_change_orders ----------
let
    Source = Csv.Document(File.Contents(DataFolder & "change_orders.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    EmptyToNull = Table.ReplaceValue(Headers, "", null, Replacer.ReplaceValue, {"approved_date"}),
    Typed = Table.TransformColumnTypes(EmptyToNull, {
        {"co_id", type text}, {"project_id", type text}, {"cost_code", type text}, {"description", type text},
        {"submitted_date", type date}, {"approved_date", type date}, {"status", type text},
        {"cost_amount", Currency.Type}, {"price_amount", Currency.Type}}, "en-US")
in
    Typed


// ---------- fact_commitments ----------
// Subcontracts and purchase orders. Line 1 is the original amount, later lines are commitment changes.
let
    Source = Csv.Document(File.Contents(DataFolder & "commitments.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Headers, {
        {"po_id", type text}, {"po_line", Int64.Type}, {"project_id", type text}, {"cost_code", type text},
        {"vendor", type text}, {"commitment_date", type date}, {"description", type text},
        {"amount", Currency.Type}}, "en-US")
in
    Typed


// ---------- fact_actuals ----------
// ERP job-cost export. Cleaning steps:
//   1. ids: trim and upper-case ("p02" -> "P02")
//   2. cost codes: trim, and put the dash back ("03300" -> "03-300")
//   3. codes that do not exist in the cost code list go to 99-999 Unallocated (shown as a data-quality KPI)
//   4. the same export batch was loaded twice: keep one row per txn_id
//   Reversals (negative amounts) are real journal entries and stay in.
let
    Source = Csv.Document(File.Contents(DataFolder & "erp_job_cost_export.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Ids = Table.TransformColumns(Headers, {{"project_id", each Text.Upper(Text.Trim(_)), type text}}),
    CodeFormat = Table.TransformColumns(Ids, {{"cost_code", each
        let t = Text.Trim(_) in
            if Text.Length(t) = 5 and not Text.Contains(t, "-") then Text.Start(t, 2) & "-" & Text.End(t, 3) else t,
        type text}}),
    ValidCodes = List.Buffer(dim_cost_code[cost_code]),
    CodeMapped = Table.TransformColumns(CodeFormat, {{"cost_code",
        each if List.Contains(ValidCodes, _) then _ else "99-999", type text}}),
    EmptyToNull = Table.ReplaceValue(CodeMapped, "", null, Replacer.ReplaceValue, {"po_id"}),
    Typed = Table.TransformColumnTypes(EmptyToNull, {
        {"txn_id", type text}, {"posting_date", type date}, {"source", type text}, {"vendor", type text},
        {"po_id", type text}, {"description", type text}, {"amount", Currency.Type}}, "en-US"),
    Deduped = Table.Distinct(Typed, {"txn_id"})
in
    Deduped


// ---------- fact_billing ----------
let
    Source = Csv.Document(File.Contents(DataFolder & "billing.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Headers, {
        {"project_id", type text}, {"billing_period_end", type date}, {"application_no", Int64.Type},
        {"billed_amount", Currency.Type}, {"retention_held", Currency.Type}}, "en-US")
in
    Typed


// ---------- fact_progress ----------
// Weekly physical % complete typed by each PM (table tblProgress). Cleaning steps:
//   1. drop blank rows, trim and upper-case ids
//   2. "45" typed instead of "45%": values above 1 are divided by 100
//   3. a week entered twice: keep the last row (the correction)
let
    Tables = Table.AddColumn(pm_workbooks, "Data", each
        Table.SelectRows([Workbook], each [Kind] = "Table" and [Item] = "tblProgress"){0}[Data]),
    Combined = Table.Combine(Tables[Data]),
    NotBlank = Table.SelectRows(Combined, each [project_id] <> null and [week_ending] <> null and [pct_complete] <> null),
    Ids = Table.TransformColumns(NotBlank, {{"project_id", each Text.Upper(Text.Trim(Text.From(_))), type text}}),
    Typed = Table.TransformColumnTypes(Ids, {{"week_ending", type date}, {"pct_complete", type number}, {"comment", type text}}, "en-US"),
    Scaled = Table.TransformColumns(Typed, {{"pct_complete", each if _ > 1 then _ / 100 else _, type number}}),
    Indexed = Table.AddIndexColumn(Scaled, "row", 1, 1, Int64.Type),
    LastPerWeek = Table.Group(Indexed, {"project_id", "week_ending"}, {
        {"pct_complete", each Table.Max(_, "row")[pct_complete], type number},
        {"comment", each Table.Max(_, "row")[comment], type text}})
in
    LastPerWeek


// ---------- fact_forecast ----------
// Monthly forecast to complete by cost code (table tblForecast), one snapshot per month end.
let
    Tables = Table.AddColumn(pm_workbooks, "Data", each
        Table.SelectRows([Workbook], each [Kind] = "Table" and [Item] = "tblForecast"){0}[Data]),
    Combined = Table.Combine(Tables[Data]),
    NotBlank = Table.SelectRows(Combined, each [project_id] <> null and [cost_code] <> null and [forecast_month] <> null),
    Ids = Table.TransformColumns(NotBlank, {
        {"project_id", each Text.Upper(Text.Trim(Text.From(_))), type text},
        {"cost_code", each Text.Trim(Text.From(_)), type text}}),
    Typed = Table.TransformColumnTypes(Ids, {
        {"forecast_month", type date}, {"forecast_to_complete", Currency.Type}, {"note", type text}}, "en-US")
in
    Typed
