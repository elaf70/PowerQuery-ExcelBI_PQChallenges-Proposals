/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7025310321896505344-gkpI
PQ proposal in 3 main steps
1) ...
2) ...
3) ... 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData048"]}[Content],
    Sort = Table.Sort(Source,{"Order No", Order.Ascending}),
    Transform = List.Accumulate(Table.ToRecords(Sort), [R={}, C=0, NC=0, RT=0],(s,c)=>let
        _C = Text.Split(c[Order No],"-"){0},
        _Q = c[Quantity],
        _NC = if (_C <> s[C]) then 1 else s[NC]+Number.From(s[RT]+_Q > c[Max Size]),
        _RT = if (_C <> s[C] or s[NC]<>_NC) then _Q else s[RT]+_Q
        in [R=s[R]& {c & [#"No of Containers"=_NC, #"Running Total"=_RT]}, C=_C, NC=_NC, RT=_RT] ),
    ToTable = Table.FromRecords(Transform[R])  
in
    ToTable


