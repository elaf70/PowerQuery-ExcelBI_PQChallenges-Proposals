/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7088729442193186816-gR2l
PQ proposal in 3 main steps
    (1) ChangeType DateTime 
    (2) Group by Process & Transform each Record to add RTotal+Duration 
    (3) Combine all Resulted Records to switch back to Table
==================================*/
let
    fxDuration.ToHours = (d as duration) =>Number.Round(Number.From(d)*24,1),

    Source = Excel.CurrentWorkbook(){[Name="tData098"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source, {{"Start Time", type datetime}, {"Finish Time", type datetime}}),
    Group = Table.Group(ChangeType, {"Process"}, {"G", (_T)=>let
     _GD = fxDuration.ToHours(List.Max(_T[Finish Time])-List.Min(_T[Start Time])), 
     _TR = List.Accumulate(Table.ToRecords(_T), [r={}, rt=0], (s,c)=>let
        _SD= fxDuration.ToHours(c[Finish Time]-c[Start Time]),
        _RT = s[rt] + _SD,
        _NewRecord = Record.Combine({c, [#"Running Total"=_RT, Duration=_GD]})
        in [r=s[r] & {_NewRecord} , rt=_RT])
     in _TR[r] }),
    ToTable = Table.FromRecords(List.Combine(Group[G]))
in
    ToTable


