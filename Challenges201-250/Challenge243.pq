/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7273547271684653056-f3Qr

PQ proposal in 2 main steps
1) Group by [Emp Id] & Transform to inner-table
2) Combine all 

==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData243"]}[Content],
  Group = Table.Group(Source, "Emp ID", {"G", (t)=>let
    _L = List.Transform(t[Group], each Text.Split(_,", ")),
    _CN = List.Transform({1..Table.RowCount(t)}, each "Group"&Text.From(_)),
    _T = Table.FromColumns({{t[Emp ID]{0}}} & _L, {"Emp ID"} & _CN)
    in Table.FillDown(_T, {"Emp ID"}) }),
  Result = Table.Combine(Group[G])
in
  Result
