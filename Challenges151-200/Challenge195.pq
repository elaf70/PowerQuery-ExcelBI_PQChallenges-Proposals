/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7212665760391843840-7CLQ

PQ proposal in 5 main steps
    ...
==================================*/
let
    Source = each Excel.CurrentWorkbook(){[Name = _]}[Content],
    T1 = let 
        _T = Table.TransformColumns(Source("tData195"), {},
            each Splitter.SplitTextByAnyDelimiter({",","&",";"})(Text.Replace(Text.From(_), " ", ""))),
        _Transform = List.Transform(Table.ToRows(_T), each  
            List.Transform(List.Zip(_),  each {_{0}, Number.From(_{1})*Number.From(_{2})}) )
        in Table.FromRows(List.Combine(_Transform) , {"Items", "Total"}),
    T2 = Table.TransformColumns(Source("tData195b"), {"Items", Splitter.SplitTextByDelimiter(", ")}),
    PurchasedItems = List.Combine(T2[Items]),
    TransformRows = Table.TransformRows(T2, each let
        _AP = List.Accumulate([Items], 0, (s,c)=>let
            _Quota = List.Count(List.Select(PurchasedItems, each _=c)),
            _Total = T1[Total]{List.PositionOf(T1[Items], c)}
            in s + (try (_Total / _Quota) otherwise 0) )
        in {[Stockist], _AP} ),
    Result = Table.FromRows(TransformRows, {"Stockist", "Amount Paid"})
in
    Result

