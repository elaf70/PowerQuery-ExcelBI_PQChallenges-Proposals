/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7225712087539286016-XgfP

PQ proposal in 2 main steps
    1) Use Table.Group ("local") to build for each subset a group of inner-subtables to combine
    2) Combine all, as final result
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData206"]}[Content],
    Group = Table.Group(Source, "Group1", 
        {"G", each let
            _Rows = Table.ToRows(Table.SelectRows(_, each [Group1]<>null)),
            _RC = List.Count(_Rows),
            _CN = List.Split(List.Transform({1..2*_RC}, each "Column"&Text.From(_)),2),
            _SubTables = List.Transform({0.._RC-1}, each Table.FromRows(List.Split(_Rows{_},2), _CN{_}) )
            in Table.Combine(_SubTables)
        }, 
        GroupKind.Local, (x,y)=>if (y=null) then 1 else 0),
    Combine = Table.Combine(Group[G])
in
    Combine


