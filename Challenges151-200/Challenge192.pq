/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7207955066543054848-2Q_9

PQ proposal in 2 main steps : 
    1) Group
    2) Expand
==================================*/
let
    fxCompare = (v1, v2, Status as list) => Status{Number.Sign(Number.From(v1)-Number.From(v2))+1},
    Source = Excel.CurrentWorkbook(){[Name="tData192"]}[Content],
    Group = Table.Group(Source, {"Project", "Phase"}, {"G", each let
        Add_NWD = Table.AddColumn(_, "Days", each let
            _Dates = List.Dates(Date.From([From Date]), Duration.Days([To Date]-[From Date])+1, #duration(1,0,0,0))
            in List.Count(List.Select(_Dates, each Date.DayOfWeek(_, Day.Monday)<5)) ),
        iPosActual = List.PositionOf(_[Column1], "Actual"), 
        _A = Add_NWD{iPosActual}, _P = Add_NWD{Number.Mod(1+iPosActual,2)} 
        in [SP=fxCompare(_A[To Date], _P[To Date], {"Underrun","On Time","Overrun"}), 
            CP=fxCompare(_A[Days], _P[Days], {"Underrun","At Cost","Overrun"})] },
        GroupKind.Local, (x,y)=>Number.From(y[Phase]<>null)),
    Expand = Table.ExpandRecordColumn(Group, "G", {"SP","CP"},{"Schedule Performance", "Cost Performance"})
in
    Expand


