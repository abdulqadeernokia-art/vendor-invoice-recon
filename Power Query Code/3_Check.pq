// One row per customer invoice line of the vendor invoices being checked.
// Which vendor invoices: those listed on sheet "Vendor Invoice Headers"; if that list is empty = ALL.
let
    Vendor     = VendorDetails,
    // history: every vendor invoice (any month) that contained each customer invoice
    History    = Table.Group(Vendor, {"Provider invoice nr"}, {{"AllVendorInv", each List.Distinct([Vendor Invoice Nr])}}),
    Wanted     = List.Buffer(List.Transform(List.RemoveNulls(Excel.CurrentWorkbook(){[Name="tblVendorInvoices"]}[Content][INVOICE]),
                    each Text.Trim(Text.From(_)))),
    ThisMonth  = if List.IsEmpty(Wanted) then Vendor else Table.SelectRows(Vendor, each List.Contains(Wanted, [Vendor Invoice Nr])),
    AddHist    = Table.ExpandTableColumn(Table.NestedJoin(ThisMonth, {"Provider invoice nr"}, History, {"Provider invoice nr"}, "H", JoinKind.LeftOuter), "H", {"AllVendorInv"}),
    AlsoIn     = Table.AddColumn(AddHist, "Also billed in", each Text.Combine(List.RemoveItems([AllVendorInv], {[Vendor Invoice Nr]}), ", "), type text),
    NewRepeat  = Table.AddColumn(AlsoIn, "New or Repeat", each if [Also billed in] = "" then "New" else "Repeat", type text),
    AddPBI     = Table.ExpandTableColumn(Table.NestedJoin(NewRepeat, {"Provider invoice nr"}, PBI_Status, {"Invoice Reference"}, "P", JoinKind.LeftOuter), "P", {"PBI Status", "PBI OK"}),
    Status     = Table.AddColumn(AddPBI, "Status", each if [PBI Status] = null then "Not in PBI - already paid" else [PBI Status], type text),
    Result     = Table.AddColumn(Status, "Result", each if [PBI Status] = null or [PBI OK] = true then "OK" else "BLOCK", type text),
    Cleaned    = Table.RemoveColumns(Result, {"AllVendorInv", "PBI Status", "PBI OK"}),
    First      = {"Key", "Vendor Invoice Nr", "Provider invoice nr", "Result", "Status", "New or Repeat", "Also billed in"},
    Ordered    = Table.ReorderColumns(Cleaned, First & List.RemoveItems(Table.ColumnNames(Cleaned), First)),
    Sorted     = Table.Sort(Ordered, {{"Vendor Invoice Nr", Order.Ascending}, {"Provider invoice nr", Order.Ascending}})
in
    Sorted