/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7017701419877052416-9yEP
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData042"]}[Content],
    Group = Table.Group(Source, {"Alphabets"}, {"Data", each let
        _Seq = List.Buffer(_[Seq]),   
        _Transform = List.Accumulate( _[Seq], [R={}, From=_Seq{0}], (s,c)=> let       
            _PrevIsInSeq = List.Contains(_Seq,c-1),
            _NextIsInSeq = List.Contains(_Seq,c+1),
            _From = if (_PrevIsInSeq=false) then c else s[From],
            _NewR = if (_NextIsInSeq=true) then s[R]
                else let
                    _PrevTo = if (List.Count(s[R])>0) then List.Last(s[R])[To]+1 else c,
                    _Missed = Text.Combine(List.Transform({_PrevTo..(_From-1)}, Text.From), ",")
                    in s[R] & {[From=_From, To=c, Missed=_Missed]}
            in [From=_From, R=_NewR] )
        in Table.FromRecords(_Transform[R]) }),
    Expand = Table.ExpandTableColumn(Group, "Data", Table.ColumnNames(Group[Data]{0}))
in
    Expand
