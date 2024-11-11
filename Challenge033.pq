/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7007192171833372672-YfR_
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData033"]}[Content],
    TransformColumns = let
        _Transform= List.Transform(Table.ToColumns(Source), each let
            _L = List.Sort(List.RemoveNulls(_)),
            _R = List.Accumulate(List.Distinct(_L), {}, (s,c)=>s & {{c, List.Count(List.Select(_L,each _=c))}} )
            in _R)
        in Table.FromRows(List.Combine(_Transform), {"Letter", "Count"}), 
    TransformRows = let 
        _Group = Table.Group(TransformColumns, {"Letter"}, {"MaxCount", each List.Max([Count]), type number}),
        _Sort = Table.Sort(_Group,{{"Letter", Order.Ascending}}),
        _Combine = List.Combine(Table.TransformRows(_Sort, each List.Repeat({_[Letter]}, _[MaxCount]))) 
        in Table.FromColumns({_Combine},{"Answer Expected"}) 
in
    TransformRows


