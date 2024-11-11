/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7005017104521310208-_QDt
PQ proposal in 8 main steps
1) ...
...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData032"]}[Content],
    Split_GFGA = Table.TransformColumnTypes(
        Table.SplitColumn(Source, "GF-GA", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"GF", "GA"}),
        {{"GF", Int64.Type}, {"GA", Int64.Type}}),
    Add_Score = Table.AddColumn(Split_GFGA, "Score", each [Won]*3+[Draw]),
    Add_GD = Table.AddColumn(Add_Score, "GD", each [GF]-[GA]),
    Group = Table.Group(Add_GD, {"Group"}, {"Data", each let
        _Sort = Table.Sort(_,{{"Score", Order.Descending},{"GD", Order.Descending},{"GF", Order.Descending}})
        in {_Sort[Team]{0}, _Sort[Team]{1}} }),
    ToTable = Table.FromRows(Group[Data], {"T1","T2"}),
    ReorderT1 = List.Accumulate({1,3,4,2,5,7,6,8}, {}, (s,c)=>s & {ToTable[T1]{c-1}}),
    ReorderT2 = List.Accumulate({2,4,3,1,6,8,5,7}, {}, (s,c)=>s & {ToTable[T2]{c-1}}),
    Result = Table.FromRows(List.Zip({ReorderT1, ReorderT2}), {"Team1", "Team2"})
in
    Result

