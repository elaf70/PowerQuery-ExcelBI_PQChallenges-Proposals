/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7022412462142468096-tDfX
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData045"]}[Content],
    NewColNames = Table.ColumnNames(Source) & {"Running Total"},
    ChangeType = Table.TransformColumnTypes(Source,{{"Date", type date}, {"Amount", Int64.Type}, {"Dealer", type text}}),
    Group = Table.Group(ChangeType, {"Dealer"}, {"Data", each let
        _SubTable = _, 
        _Rows = List.Accumulate(Table.ToRows(_SubTable), 
            [R={}, DStart=Date.StartOfMonth(List.Min(_SubTable[Date])), RT=0],
            (s,c)=>let
                _WithinPeriod = Number.From(c{1}<=Date.AddMonths(s[DStart],3)),
                _DS = if (_WithinPeriod=1) then s[DStart] else c{1},
                _RT = if (_WithinPeriod=1) then s[RT]+c{2} else c{2}
                in [R=s[R] & {c & {_RT}}, DStart=_DS, RT=_RT] )[R]
        in Table.FromRows(_Rows, NewColNames) }),
    Expand = Table.ExpandTableColumn(Group[[Data]], "Data", NewColNames)
in
    Expand


