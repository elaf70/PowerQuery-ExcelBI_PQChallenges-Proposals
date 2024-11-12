/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7147073892300369920-kipN

PQ proposal in 2 main steps
    1) Group by [Group] to create 2 inner tables for each, with a new RunningTotal column 
    2) Combine all inner tables
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData144"]}[Content],
    Group = Table.Group(Source,"Group", {"All", (t)=>let
        _V = t[Value],
        N=Table.RowCount(t), 
        RT = List.Transform(List.Split(_V, Number.RoundUp(N/2)), 
              each List.Skip(List.Accumulate(_, {0}, (s,c)=>s&{List.Last(s)+c})) )
        in Table.FromColumns({_V, List.Combine(RT)}, {"Value", "Running Total"}) }),
    Expand = Table.ExpandTableColumn(Group, "All", Table.ColumnNames(Group[All]{0})   )
in
    Expand


