/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7058288870244651008-uy72
PQ proposal in 4 main steps
    (1) Add Index 
    (2) Group and Combine from Ranking+ReplaceValue 
    (3) Resort from Index 
    (4) Remove Index
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData074"]}[Content],
    Add_Index = Table.AddIndexColumn(Source, "Index", 0, 1, Int64.Type),
    Group = Table.Group(Add_Index, {"Name"}, {"Data", each let
        _nID = List.Count(List.Distinct(_[Emp ID])),
        _AddRank = Table.AddRankColumn(_, "Answer", "Emp ID", [RankKind = RankKind.Dense]),
        _Result = Table.ReplaceValue(_AddRank, null, each [Name], 
            (curr,cond,repl)=>repl & (if (_nID=1) then "" else " " & Character.FromNumber(64+curr)), {"Answer"})
        in _Result }),
    Sort = Table.Sort(Table.Combine(Group[Data]), {"Index", Order.Ascending}),
    RemoveIndex = Table.RemoveColumns(Sort,{"Index"})
in
    RemoveIndex


