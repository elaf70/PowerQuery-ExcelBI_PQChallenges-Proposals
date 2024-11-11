/*================================
Ref = https://www.linkedin.com/posts/excelbi_excel-excelchallenge-powerquerychallenge-activity-6991971220984344576-q4Ej
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData021_1"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source, {"Date", type date}),
    Sort = Table.Sort(ChangeType,{{"Name", Order.Ascending}, {"Date", Order.Ascending}}),
    Group = Table.Group(Sort, {"Name", "Won"}, 
        {"Date", each let 
           _LD = List.Buffer(_[Date]),
           _RD = List.Accumulate(_LD, {}, (s,c)=> let
               _NBOK = List.Count(List.Select(_LD, each Abs(Duration.Days(c - _))<=180))
                in s & (if (_NBOK>=3) then {c} else {}) )
           in _RD} ,
        GroupKind.Local ),
    Filter = Table.SelectRows(Group, each ([Won] = "Y" and List.Count([Date])>=3))[[Name],[Date]],
    Expand = Table.ExpandListColumn(Filter, "Date")
in
    Expand

