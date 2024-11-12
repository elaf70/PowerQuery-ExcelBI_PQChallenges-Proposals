/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7109023163644076032-1U65

PQ proposal in 2 main steps
    1) Group
    2) Expand
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData114"]}[Content],
    T1 = Excel.CurrentWorkbook(){[Name="tData114_1"]}[Content],
    CN = {"Early","On Time","Late","Out of Limit"},
    Group = Table.Group(Source, {"Group"}, {"Data", each let
        _G =  _[Group]{0},
        _TG = Table.SelectRows(T1, each [Group]=_G),
        _LS = List.Accumulate(_[Start Time], {}, (s,_TStart)=> let
            _AddStatus = Table.AddColumn(_TG, "Status", each let
                st=[Start Time], et=[End Time], st30=st-1/48, et30=et+1/48
                in 
                if (_TStart>=st and _TStart<=et) then "On Time"
                else (if ((st30>0 and _TStart<st and _TStart>=st30) or (st30<0 and (_TStart>1+st30 or _TStart<st))) then "Early"
                else (if ((et30<1 and _TStart>et and _TStart<=et30) or (et30>1 and (_TStart<1-et30 or _TStart>et))) then "Late"
                else "Out of Limit")) ),
            _S = if (       Table.RowCount(Table.SelectRows(_AddStatus, each [Status]="On Time"))>0) then "On Time"
                else (if (  Table.RowCount(Table.SelectRows(_AddStatus, each [Status]="Late"  ))>0) then "Late"
                else (if (  Table.RowCount(Table.SelectRows(_AddStatus, each [Status]="Early" ))>0) then "Early"
                else "Out of Limit"))
            in s & {_S}),
          _T = Table.FromColumns({_LS}, {"Status"})  
        in Table.Pivot(_T, List.Distinct(_T[Status]), "Status", "Status", List.Count)  }),
    Expand = Table.ExpandTableColumn(Group, "Data", CN)
in
    Expand


