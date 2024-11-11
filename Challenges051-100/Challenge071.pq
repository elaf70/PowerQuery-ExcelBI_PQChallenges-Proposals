/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7055389763721175040-SGBT
PQ proposal in 2 main steps
    1) Group
    2) Combine
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData071"]}[Content],
    NewColNames= Table.ColumnNames(Source) & {"Running total"},
    Group = Table.Group(Source, {"Group"}, {"Data", each let
        _RT = List.Accumulate([Weight], [RT={}, R=0], (s,c)=>let
            _New = c + s[R],     
            _RT = if (_New<=100) then _New else 100,
            _R = if (_New<=100) then _New else _New-100
            in [RT=s[RT]&{_RT}, R=_R])[RT]
        in Table.FromColumns(Table.ToColumns(_)&{_RT}, NewColNames) }),
    Combine = Table.Combine(Group[Data])
in
    Combine

