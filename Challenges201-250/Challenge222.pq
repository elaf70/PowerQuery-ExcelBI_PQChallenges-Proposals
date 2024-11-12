/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7246005780800299010-qtmD

PQ proposal in 4 main steps
    1) Transform each row in a table 
    2) Group by Subject to build sub-tables of sorted students who passed
    3) Expand all these sub-tables 
    4) Sort by Subject 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData222"]}[Content],
    N = (Table.ColumnCount(Source)-1) / 2, 
    Transform = Table.TransformRows(Source, each let 
        _v = Record.FieldValues(_),
        _lsm = List.Select(List.Zip(List.Split(List.Skip(_v),N)), each _{0}<>null) ,
        _t = Table.FromRows(_lsm, {"Subject", "Mark"})
        in Table.AddColumn(_t, "Student", each _v{0}) ),
    Group = Table.Group(Table.Combine(Transform), "Subject", {"G", each let
        _v = List.Sort(Table.SelectRows(_, each [Mark]>=40)[Student]),
        _CN = List.Transform({1..List.Count(_v)}, each "Studend_" & Text.From(_)),
        _Result = Table.FromRows({_v}, _CN) 
        in _Result}),    
    Expand = Table.ExpandTableColumn(Group, "G", 
        List.Max(List.Transform(Group[G], Table.ColumnNames),{}, (x,y)=>Number.From(List.Count(x)>List.Count(y)))),
    Sort = Table.Sort(Expand,{{"Subject", Order.Ascending}})
in 
    Sort



