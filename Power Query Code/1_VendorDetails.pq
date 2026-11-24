// Vendor invoice details straight from your PDF Converter workbook (no copy-paste).
// File location: sheet Settings, row "Data Folder" (default = this workbook's folder).
// To change the table/sheet: click the gear next to step "Navigation".
let
    Settings   = Excel.CurrentWorkbook(){[Name="tblSettings"]}[Content],
    Folder0    = Text.Trim(Text.From(Settings{[Setting="Data Folder"]}[Value])),
    Folder     = if Text.EndsWith(Folder0, "\") then Folder0 else Folder0 & "\",
    Source     = Excel.Workbook(File.Contents(Folder & "Sample Inputs\PDF Converter.xlsx"), null, true),
    Navigation = Source{[Item="InvoiceDetails",Kind="Table"]}[Data],

    // column names of YOUR file are typed on sheet Settings
    VendCol    = Text.Trim(Text.From(Settings{[Setting="Vendor Invoice Column"]}[Value])),
    CustCol    = Text.Trim(Text.From(Settings{[Setting="Customer Invoice Column"]}[Value])),

    Renamed    = Table.RenameColumns(Navigation, List.Select({{VendCol, "Vendor Invoice Nr"}, {CustCol, "Provider invoice nr"}}, each _{0} <> _{1})),
    // invoice numbers as clean text so 22222 and "22222" match
    AsText     = Table.TransformColumns(Renamed, {{"Vendor Invoice Nr", each if _ = null then null else Text.Trim(Text.From(_)), type text},
                                                  {"Provider invoice nr", each if _ = null then null else Text.Trim(Text.From(_)), type text}}),
    NoBlanks   = Table.SelectRows(AsText, each [Provider invoice nr] <> null and [Provider invoice nr] <> ""),
    // KEY linking vendor invoice and customer invoice
    AddKey     = Table.AddColumn(NoBlanks, "Key", each [Vendor Invoice Nr] & "|" & [Provider invoice nr], type text),
    // any column with "Date" in its name becomes a real date (bad values -> empty)
    DateCols   = List.Select(Table.ColumnNames(AddKey), each Text.Contains(Text.Lower(_), "date")),
    Dates      = Table.TransformColumns(AddKey, List.Transform(DateCols, (c) => {c, each try Date.From(_) otherwise null, type nullable date})),
    First      = {"Key", "Vendor Invoice Nr", "Provider invoice nr"},
    Ordered    = Table.ReorderColumns(Dates, First & List.RemoveItems(Table.ColumnNames(Dates), First))
in
    Ordered