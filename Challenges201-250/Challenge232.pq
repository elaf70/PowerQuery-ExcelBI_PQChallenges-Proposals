/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7258689352111702016-JKZz

PQ proposal in 2 main steps
    1) Group per Store to create inner-tables per date with List.TransformMany
    to get all dates + cumul of qty between each original date & next one
    2) Combine all tables
==================================*/
let
    Source = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name="tData232"]}[Content], {"Date", type date}),
    CN = Table.ColumnNames(Source),
    Group = Table.Group(Source, {"Store"}, {"G", each let 
        _Rows = List.Buffer(Table.ToRows(_)),
        _Transform = List.TransformMany(_Rows, 
            each List.Dates(_{1}, try Duration.Days(_Rows{List.PositionOf(_Rows,_)+1}{1}-_{1}) otherwise 1, #duration(1,0,0,0)),
            (x,_) => {x{0}, _, x{2}*(Duration.Days(_-x{1})+1)} )
        in Table.FromRows(_Transform, CN) }),
    Combine = Table.Combine(Group[G])
in 
    Combine


