/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7002117425731112960-5Nx4
PQ proposal in 6 main steps
1) ...
2) ...
3) ... 
4) ...
5) ... 
6) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData29"]}[Content],
    Unpivot = Table.UnpivotOtherColumns(Source, {"Year"}, "Attribute", "Team")[[Team],[Year]],
    Group = Table.Group(Unpivot, {"Team"}, 
        {   {"Years", each List.Sort(_[Year])}, {"Count", each List.Count(_[Year])} }),
    Filter = Table.SelectRows(Group, each ([Count] <> 1)),
    Transform = Table.TransformColumns(Filter[[Team], [Years]], 
        {"Years", each let
            _Y = List.Sort(_),
            _LYNY = List.Zip({List.RemoveLastN(_Y,1), List.Skip(_Y)}),   // List : {{Year, NextYear}...}
            _T = Table.FromRows(_LYNY, {"Y", "NY"} ),
            _G = Table.AddColumn(_T, "Gap", each [NY]-[Y]),
            _MaxG = List.Max(_G[Gap]),
            _Filter = Table.SelectRows(_G, each [Gap]=_MaxG),
            _Years = Table.AddColumn(_Filter, "Years", each Text.From([Y])&"-"&Text.From([NY])),
            _GY = Text.Combine(_Years[Years], ", ")   
            in [MaxGap=_MaxG, Years=_GY] } ),
    Expand = Table.ExpandRecordColumn(Transform, "Years", {"MaxGap", "Years"}),
    Sort = Table.Sort(Expand,{{"MaxGap", Order.Descending}, {"Team", Order.Ascending}})
in
    Sort
