/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7174615372044656640-nvC_

PQ proposal in 3 main steps
    1) Group by Dept and add [Total] row on each 
    2) Add [Grand Total] row
    3) Add Bons column
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData165"]}[Content],
    CN = Table.ColumnNames(Source),
    fxTTotal = (t, sTotal)=> #table(CN, {{sTotal, List.Count(t[Emp]), List.Sum(t[Salary])}}),
    Group = Table.Group(Source, "Dept", {"All", each _ & fxTTotal(_, "Total")},0, (p,c)=> Number.From(c<>null)),
    Add_GT = Table.Combine(Group[All]) & fxTTotal(Source, "Grand Total"),
    Add_B = Table.AddColumn(Add_GT, "Max Bonus", each [Salary]*0.1)
in
    Add_B


