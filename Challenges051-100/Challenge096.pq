/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7086192726538358784-yTdB
PQ proposal in 3 main steps
    1) Join
    2) Expand
    3) Filter
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData096"]}[Content],
    CNames = Table.ColumnNames(Source),
    CType = Table.TransformColumnTypes(Source,
        {{"Store1 Visit Date Time", type datetime}, {"Store2 Visit Date Time", type datetime}}),
    Join = Table.NestedJoin(
        Table.SelectColumns(CType, List.FirstN(CNames,2)), CNames{0}, 
        Table.SelectColumns(CType, List.LastN (CNames,2)), CNames{2},
        "T2", JoinKind.LeftOuter),
    Expand = Table.ExpandTableColumn(Join, "T2", {"Store2 Visit Date Time"}),
    Filter = Table.SelectRows(Expand, each [Store1 Visit Date Time]<=[Store2 Visit Date Time])
in
    Filter


