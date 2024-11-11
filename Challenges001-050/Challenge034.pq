/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7007554559728316416-WUNA
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData034"]}[Content],
    ColumnsToList = Table.FromColumns(
        List.Transform(Table.ToColumns(Source), each {List.RemoveNulls(_)}),
        Table.ColumnNames(Source) ),
    Expand = List.Accumulate({"Country", "Manager","Region"}, ColumnsToList,        (s,c)=>Table.ExpandListColumn(s, c) )
in 
    Expand

