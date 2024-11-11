/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6989434734816714752-Iys4
PQ proposal in 4 main steps
1) ...
2) ...
3) ...
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData19"]}[Content],
    Subjects = List.Skip(Table.ColumnNames(Source)),
    GetLists = List.Accumulate({1..3}, {}, 
        (s1, c1) => s1 & { {{c1}} & 
            List.Accumulate(Subjects, {}, 
                (s2, c2) => s2 & {Table.SelectRows( 
                    Table.AddRankColumn( Source, c2 & "_rank", {c2, Order.Descending}, [RankKind = RankKind.Dense]), 
                    each Record.Field(_, c2 & "_rank") = c1)[#"Name\Subject"]})}),
    Combine = Table.Combine(List.Transform(GetLists, each Table.FromColumns(_, {"Rank"} & Subjects))),
    FillDown = Table.FillDown(Combine,{"Rank"})
in
    FillDown

