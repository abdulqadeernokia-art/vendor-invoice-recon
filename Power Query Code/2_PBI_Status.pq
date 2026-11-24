// One status per customer invoice, built from query PBI_Data (sheet "PBI Data").
// Change the Power BI file in query PBI_Data, not here.
let
    KeepCols   = Table.SelectColumns(PBI_Data, {"Invoice Reference", "Line Item Status Check"}),
    AsText     = Table.TransformColumns(KeepCols, {{"Invoice Reference", each if _ = null then null else Text.Trim(Text.From(_)), type text},
                                                  {"Line Item Status Check", each if _ = null then null else Text.From(_), type text}}),
    // statuses that allow approval, from sheet Status Rules
    Rules      = Excel.CurrentWorkbook(){[Name="tblStatusRules"]}[Content],
    OkPrefixes = List.Buffer(List.Transform(Table.SelectRows(Rules, each [#"Can approve?"] = "Yes")[#"Status starts with"], Text.Lower)),
    LineOK     = Table.AddColumn(AsText, "Line OK", (r) =>
                    List.AnyTrue(List.Transform(OkPrefixes, (p) => Text.StartsWith(Text.Lower(r[Line Item Status Check] ?? ""), p))),
                    type logical),
    // several lines per invoice: OK only if ALL lines are OK
    Grouped    = Table.Group(LineOK, {"Invoice Reference"}, {
                    {"PBI Status", each Text.Combine(List.Distinct(List.RemoveNulls([Line Item Status Check])), "; "), type text},
                    {"PBI OK",     each List.AllTrue([Line OK]), type logical}})
in
    Grouped