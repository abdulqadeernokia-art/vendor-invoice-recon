// ONE ROW PER VENDOR INVOICE. Built from the Check query - nothing to paste here.
let
    C          = Check,
    Settings   = Excel.CurrentWorkbook(){[Name="tblSettings"]}[Content],
    FeeCol     = Text.From(Settings{[Setting="Fee Column Name"]}[Value]),
    HasFee     = List.Contains(Table.ColumnNames(C), FeeCol),
    CountInv   = (t as table) => List.Count(List.Distinct(t[Provider invoice nr])),
    Grouped    = Table.Group(C, {"Vendor Invoice Nr"}, {
        {"Customer Invoices",    each CountInv(_), Int64.Type},
        {"Detail Lines",      each Table.RowCount(_), Int64.Type},
        {"Blocked",           each CountInv(Table.SelectRows(_, (r) => r[Result] = "BLOCK")), Int64.Type},
        {"Repeat",            each CountInv(Table.SelectRows(_, (r) => r[New or Repeat] = "Repeat")), Int64.Type},
        {"Paid (not in PBI)", each CountInv(Table.SelectRows(_, (r) => Text.StartsWith(r[Status], "Not in PBI"))), Int64.Type},
        {"Fee in Details",    each if HasFee then Number.Round(List.Sum(List.Transform(Table.Column(_, FeeCol), (v) => try Number.From(v) otherwise 0)), 2) else null, type nullable number}}),
    // optional: vendor invoice amounts pasted on sheet "Vendor Invoice Headers"
    Hdr0       = Excel.CurrentWorkbook(){[Name="tblVendorInvoices"]}[Content],
    Hdr1       = Table.SelectRows(Hdr0, each [INVOICE] <> null),
    Hdr        = Table.Group(Table.TransformColumns(Hdr1, {{"INVOICE", each Text.Trim(Text.From(_)), type text}}),
                    {"INVOICE"}, {{"Vendor Invoice Amount", each List.Sum(List.Transform([AMOUNT], (v) =>
                        try Number.From(v) otherwise try Number.From(v, "en-US") otherwise null)), type nullable number}}),
    Joined     = Table.ExpandTableColumn(Table.NestedJoin(Grouped, {"Vendor Invoice Nr"}, Hdr, {"INVOICE"}, "A", JoinKind.LeftOuter), "A", {"Vendor Invoice Amount"}),
    Diff       = Table.AddColumn(Joined, "Difference", each if [Vendor Invoice Amount] = null or [Fee in Details] = null then null
                    else Number.Round([Vendor Invoice Amount] - [Fee in Details], 2), type nullable number),
    Decision   = Table.AddColumn(Diff, "Decision", each
                    if [Blocked] > 0 then "HOLD"
                    else if [Difference] <> null and Number.Abs([Difference]) > 0.01 then "CHECK AMOUNT"
                    else "APPROVE", type text),
    Sorted     = Table.Sort(Decision, {{"Vendor Invoice Nr", Order.Ascending}})
in
    Sorted