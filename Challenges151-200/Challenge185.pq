/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7199994361248428033-8QhI

PQ proposal in 2 main steps
    1) Group
    2) Combine
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData185"]}[Content],
    Group = Table.Group(Source, {"Group "}, {"G", each let
        _E = List.Buffer(List.Distinct([Emp]))
        in Table.AddColumn(_, "Index", each List.PositionOf(_E, [Emp])+1) }),
    Combine = Table.Combine(Group[G])
in
    Combine


