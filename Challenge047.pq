/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7024949177134505984-fx62
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData047"]}[Content],
    P_All = List.Sort(List.Distinct(Source[Process])),
    Group = Table.Group(Source, {"User"}, {"P", each _[Process]})[P],
    Transform = List.Accumulate(P_All, [P={}, R={}], (s,c)=>let
        _P= s[P] & {c},
        _Select = List.Select(Group, (l)=>List.ContainsAll(l, _P))
         in [P=_P, R=s[R] & {{Text.Combine(_P, ", "), List.Count(_Select)}}] )[R],
    ToTable = Table.FromRows(Transform, {"Process Group", "Count"})
in
    ToTable


