/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7121344353544941568-T_fc

PQ proposal in 2 main steps
    1) Transform each row with List.Generate to split it in a new table. 
    For each Start-Finish as required with calc of related duration
    2) Combine all of these sub-tables 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData123"]}[Content],
    Transform = Table.TransformRows(Source, each let
        _R = _,
        _L = List.Generate( ()=>[s=_R[Start], r=[]], each [s]<=_R[Finish],
            each let 
            _NextW = DateTime.From(DateTime.Date([s])) + #duration(7 - Date.DayOfWeek([s], Day.Monday),6,0,0),
            _NextS = if (_NextW < _R[Finish]) then _NextW else _R[Finish],
            _D = Duration.TotalHours(_NextS-[s]),
            _NewR = [Machine=_R[Machine], Start=[s], Finish=_NextS, Duration=_D]
            in [s=if ([s]<>_R[Finish]) then _NextS else _R[Finish]+#duration(0,0,0,1), r=_NewR],
            each [r] )
        in Table.FromRecords(List.Skip(_L)) ),
    Combine = Table.Combine(Transform)
in 
    Combine


