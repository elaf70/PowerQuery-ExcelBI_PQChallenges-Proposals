/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7230785486896734210-RO2I

PQ proposal in 3 main steps
    1) Filter rows to exlude WE dates
    2) GroupBy (Name,Type) with GroupKindLocal to get "from" +"to" dates of each subset
    3) Reorder columns as expected 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData210"]}[Content],
    Filter = Table.SelectRows(Source, each Date.DayOfWeek([Date], Day.Monday)<5),
    Group = Table.Group(Filter, {"Name","Type"},
        { {"From Date", each List.Min([Date])}, {"To Date", each List.Max([Date])} }, 
        GroupKind.Local ),
    ReorderCols = Group[[Name],[From Date],[To Date],[Type]]
in
    ReorderCols


