/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7057926472740573184-RQUw
PQ proposal in 3 main steps
    1) ...
    2) ...
    3) ... 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData073"]}[Content],
    CN=List.Skip(Table.ColumnNames(Source),2),
    ReplaceValue = Table.ReplaceValue(Source, null, each [Factor], (curr, cond, repl)=>curr*repl, CN),
    Remove_Factor = Table.RemoveColumns(ReplaceValue, {"Factor"}),
    Add_Total= Table.AddColumn(Remove_Factor, "Total", each List.Sum(List.Skip(Record.ToList(_))))
in
    Add_Total


