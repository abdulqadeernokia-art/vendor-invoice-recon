// Your full Power BI report, loaded to sheet "PBI Data" (all columns).
// File location: sheet Settings, row "Data Folder" (default = this workbook's folder).
// To change the sheet: click the gear next to step "Navigation".
let
    Settings   = Excel.CurrentWorkbook(){[Name="tblSettings"]}[Content],
    Folder0    = Text.Trim(Text.From(Settings{[Setting="Data Folder"]}[Value])),
    Folder     = if Text.EndsWith(Folder0, "\") then Folder0 else Folder0 & "\",
    Source     = Excel.Workbook(File.Contents(Folder & "Sample Inputs\PBI Status.xlsx"), null, true),
    Navigation = Source{[Item="Data",Kind="Sheet"]}[Data],
    Headers    = Table.PromoteHeaders(Navigation, [PromoteAllScalars = true]),
    NoBlanks   = Table.SelectRows(Headers, each [Invoice Reference] <> null and Text.Trim(Text.From([Invoice Reference])) <> "")
in
    NoBlanks