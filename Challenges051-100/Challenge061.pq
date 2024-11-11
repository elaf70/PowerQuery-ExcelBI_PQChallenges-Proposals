/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7042706184721567744-m8g3
PQ proposal in 4 main steps
    1) ...
    2) ...
    3) ... 
    4) ...
==================================*/

// SOLUTION1 : Without PIVOT
//=============================================
let
    Source1 = Excel.CurrentWorkbook(){[Name="tData061_1"]}[Content],
    Source2 = Excel.CurrentWorkbook(){[Name="tData061_2"]}[Content],
    Join = Table.Join(Source1, "Data", Source2, "Data"),
    Categories=List.Distinct(Source2[Category]),
    Group = Table.Group(Join, {"ID"}, {{"Data", each let
        _G = Table.Group(_, {"Category"}, {{"Data", each Text.Combine(_[Data],", ")}}) 
        in Record.FromList(_G[Data],_G[Category])} }),    
    Expand = Table.ExpandRecordColumn(Group, "Data", Categories, Categories) 
in
    Expand

/*
// SOLUTION2 : With PIVOT
//=============================================
let
    Source1 = Excel.CurrentWorkbook(){[Name="tData061_1"]}[Content],
    Source2 = Excel.CurrentWorkbook(){[Name="tData061_2"]}[Content],
    Join = Table.Join(Source1, "Data", Source2, "Data"),
    Group = Table.Group(Join, {"ID", "Category"}, {{"Data", each Text.Combine(List.Sort(_[Data]),", ")}}),
    Pivot = Table.Pivot(Group, List.Distinct(Group[Category]), "Category", "Data")
in
    Pivot
*/