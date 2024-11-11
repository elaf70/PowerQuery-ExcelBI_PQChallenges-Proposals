/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6984360699879677952-nsSW
PQ proposal in 3 main steps
1) ...
2) ...
3) ... 
==================================*/
let
    Source = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData015"]}[Content]),
    Subjects = List.Sort(List.RemoveFirstN(Table.ColumnNames(Source),1)),
    ResultValues = List.Transform( Subjects, each Table.MaxN( Source, _, 3)[#"Name\Subject"] ),
    Result = Table.FromColumns( {{1,2,3}} & ResultValues, {"Rank"} & Subjects)
in
    Result


