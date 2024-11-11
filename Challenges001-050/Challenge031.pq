/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7004655066875883520-mLy5
PQ proposal in 5 main steps
1) ...
2) ...
3) ... 
4) ...
5) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData031"]}[Content],
    Add_LastName = Table.AddColumn(Source, "LastName", each Text.AfterDelimiter([Name], " "), type text),
    Join = Table.NestedJoin(Add_LastName, {"LastName", "Country"}, Add_LastName, {"LastName", "Country"}, "Son", JoinKind.Inner),
    Expand = Table.ExpandTableColumn(Join, "Son", {"Name", "Year"}, {"Son", "Son.Year"}),
    SelectRows = Table.SelectRows(Expand, each [Son.Year]>=([Year]+16)),
    SelectCols = Table.SelectColumns(SelectRows,{"Name", "Son", "Country"})
in
    SelectCols


