/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6986899936902565888-Lgj4
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/

// SOLUTION2 : GROUP ALL DATA-TRANSFORMATIONS WITH ONLY 2 CALLS OF : LIST.ACCUMULATE
//----------------------------------------------------------------------------------
let
    Source = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData17"]}[Content]),
    Group = Table.Group(Source, {"Name"}, {{"Data", each _[Task Dates]}}),
    Transform = Table.TransformColumns(Group, {{"Data", each let
        _List=List.Buffer(_),
        _Accumulate = List.Accumulate(_List, {[From=_List{0},To=_List{0}]},
             (s,c)=>if (Duration.Days(c - List.Last(s)[To])<=1)
                    then List.RemoveLastN(s,1) & {[From=List.Last(s)[From],To=c]}
                    else s & {[From=c,To=c]}),
        _Result = List.Accumulate(_Accumulate, {},
            (s,c)=>if (Duration.Days(c[To]-c[From])>0) 
                   then s & {Record.AddField(c, "Task", "Task" & Number.ToText(List.Count(s)+1))} 
                   else s )
        in Table.FromList(_Result, Record.FieldValues, {"From", "To", "Task"})
        } }),
    Expand = Table.ExpandTableColumn(Transform, "Data", {"Task", "From", "To" }, {"Task", "From", "To"}),
    ChangeType = Table.TransformColumnTypes(Expand,{{"From", type date}, {"To", type date}})
in  ChangeType

/*
// SOLUTION 1 : With 1 call of List.accumulate + extra "more classic" table data manipulation
//-------------------------------------------------------------------------------------------
let
    Source = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData17"]}[Content]),
    Group = Table.Group(Source, {"Name"}, {{"Data", each _[Task Dates]}}),
    Transform = Table.TransformColumns(Group, {{"Data", each let
        _List=List.Buffer(_),
        _Accumulate = List.Accumulate(_List, {[From=_List{0},To=_List{0}]},
             (s,c)=>if (Duration.Days(c - List.Last(s)[To])<=1)
                    then List.RemoveLastN(s,1) & {[From=List.Last(s)[From],To=c]}
                    else s & {[From=c,To=c]}),
         _ToTable = Table.FromList(_Accumulate, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
         _Expand = Table.ExpandRecordColumn(_ToTable, "Column1", {"From", "To"}, {"From", "To"}),
         _Filter = Table.SelectRows(_Expand, each Duration.Days([To]-[From])>0),
         _Index = Table.AddIndexColumn(_Filter, "Index", 1, 1, Int64.Type),
         _Result = Table.AddColumn(_Index, "Task", each "Task" & Number.ToText([Index]))
        in _Result
        } }),
    Expand = Table.ExpandTableColumn(Transform, "Data", {"Task", "From", "To"}, {"Task", "From", "To"}),
    ChangeType = Table.TransformColumnTypes(Expand,{{"From", type date}, {"To", type date}})
in  ChangeType
*/