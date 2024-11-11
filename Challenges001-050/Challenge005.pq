/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6971677332868591616-0iGQ
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/

let
    Source = Excel.CurrentWorkbook(){[Name="tData005"]}[Content],
    CN = {"Rank By Population", "Country","2022 Population","Growth Rate","Area","Density (km2)"},
    Split = Table.FromRows(List.Split(Source[Data],5), List.Skip(CN)),
    AddRank = Table.AddRankColumn(Split, CN{0}, {CN{2},  Order.Descending}),
    ReorderCols = Table.ReorderColumns(AddRank, CN)
in
    ReorderCols

