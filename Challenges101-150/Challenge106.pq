/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7098876302987816961-GaW3

PQ proposal in 4 main steps
    1) Transform each Row in a sub-table + Find TOP3 and sort
    2) Combine all 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData106"]}[Content],
    N = (Table.ColumnCount(Source) -1) / List.Count(List.Select(Table.ColumnNames(Source), each Text.StartsWith(_, "Student"))),
    CN = List.FirstN(Table.ColumnNames(Source), N+1),
    Transform = Table.TransformRows(Source, each let
        _LV = Record.FieldValues(_),
        _SF3 = List.Sum(List.FirstN(List.Sort(_LV),3)),
        _NewRows = List.Transform(List.Split(List.Skip(_LV), N), each let
            _SF3 = List.Sum(List.LastN(List.Sort(List.Skip(_)), 3))
            in {_LV{0}} & _ & {_SF3} ),
        _T = Table.FromRows(List.Select(_NewRows, each _{1} <> null), CN & {"Best of 3 Total"}),
        _AddRank = Table.AddRankColumn(Table.Sort(_T, {"Student"}), "Rank", {"Best of 3 Total", Order.Descending})
        in Table.Sort(_AddRank, {{"Rank", Order.Ascending}, {"Student", Order.Ascending}}) ), 
    Combine = Table.Combine(Transform)
in 
    Combine


