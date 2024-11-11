/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7078220194535358464-H-LQ
PQ proposal in 4 main steps
    1) ...
    2) ...
    3) ... 
    4) ...
==================================*/
let
    STimes = let 
        _T = Excel.CurrentWorkbook(){[Name="tData089_Start"]}[Content],
        _Transform = Table.RemoveLastN(Table.UnpivotOtherColumns(_T, {}, "Start", "Hour"))
        in Table.Buffer(Table.TransformColumnTypes(_Transform,{{"Hour", type time}})),

    Source = Excel.CurrentWorkbook(){[Name="tData089"]}[Content],
    MergeP = Table.CombineColumnsToRecord(Source, "Period", {"Start Time", "End Time"}),
    Tranf_P = Table.TransformColumns(MergeP, {"Period", (r)=> let
        _TS = r[Start Time], _TE = r[End Time],
        _Hours = let
            _L1 = List.Generate(()=>_TS, each _<_TE, each _+#duration(0,0,60-Time.Minute(_),0) )
            in  List.Skip(_L1) & {_TE},
        _Periods = List.Accumulate(_Hours, [ts=_TS, i=-1, r={}], (s,c)=>let
            _i = List.Count(List.Select(STimes[Hour], each _<DateTime.Time(c)))-1,
            _New = if (_i=s[i]) then {} else {[DS=s[ts], DE=c, N=STimes[Start]{_i}]}
            in [ts=c, i=_i, r=s[r] & _New] )
        in Table.FromRecords(_Periods[r], type table[DS=DateTime.Type , DE=DateTime.Type, N=text]) }),
    Expand = Table.ExpandTableColumn(Tranf_P, "Period", {"DS", "DE", "N"}, {"Start Time", "End Time", "Start"}),
    Reorder = Table.ReorderColumns(Expand,{"Fleet", "Start Time", "End Time", "Group", "Start"})
in
    Reorder


