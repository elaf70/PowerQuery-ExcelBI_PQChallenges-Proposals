/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7253615921712259072-OyOS

PQ proposal in 7 main steps ... 
     Not that short & simple

But as usual, I was searching no to use any "Transpose".
Proper sorting was for me, by far the hardest & longuest part to find.
But instructive : as I had to search in Table.Sort doc & samples ... 😅
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData228"]}[Content],
    H2RFill = List.Transform( Table.ToRows(Table.FirstN(Source,2)), each List.Accumulate(_, {}, (s,c)=>s & {c??List.Last(s)})),
    NewH = List.Transform(List.Zip(H2RFill), each Text.Combine(_,";")),
    RenameCols = Table.RenameColumns(Table.Skip(Source,2), List.Zip({Table.ColumnNames(Source), NewH})),
    Unpivot = Table.UnpivotOtherColumns(RenameCols, {"Student"}, "Attribute", "Marks"),
    SplitAttribute = Table.SplitColumn(Unpivot, "Attribute", Splitter.SplitTextByDelimiter(";"), {"Category", "Value"}),
    SOrders = List.Buffer({H2RFill{0}, Source[Column1], H2RFill{1}}),
    Sort = Table.Sort(SplitAttribute, {
        each List.PositionOf(SOrders{0},[Category]), 
        each List.PositionOf(SOrders{1},[Student]), 
        each List.PositionOf(SOrders{2},[Value])           })
in
    Sort


