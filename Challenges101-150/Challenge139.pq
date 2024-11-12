/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7141638071652958208-ephZ

PQ proposal in 3 main steps
    1) Remowe WeekEnd
    2) Group
    3) Expand
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData139"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Vacation Date", type date}}),
    RemoveWE = Table.SelectRows(ChangeType, each Date.DayOfWeek([Vacation Date],Day.Monday)<=4),
    Group =Table.Group(RemoveWE, {"Name"}, 
        {"All", each let
            _t = List.Accumulate(_[Vacation Date], {[VN=0, Start=null, End=#date(1900,1,1), WD=0]}, 
                (s,c)=>let
                    _CurV = List.Last(s)
                    in
                        if (Duration.Days(c-_CurV[End]) > 3) 
                        then s & {[VN=_CurV[VN]+1, Start=c, End=c, WD=1]}
                        else List.RemoveLastN(s) & {Record.TransformFields(_CurV, {{"End", each c}, {"WD", each _+1}})}
                )
            in Table.FromRecords(List.Skip(_t))
        } ),
    Expand = Table.ExpandTableColumn(Group, "All", {"VN", "Start", "End", "WD"}, 
        {"Vacation No", "Vacaction From Date", "Vacation End Date", "Number of WorkDays"})
in
    Expand


