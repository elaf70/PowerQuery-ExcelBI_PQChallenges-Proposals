/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7225350748308480001-W_Q7

PQ proposal in 4 main steps
    1) Join the 2 original table
    2) Group by "Yes/no" column & Transform each as requested
    3) Combine & Sort
==================================*/
let
    Sources = Table.SelectRows(Excel.CurrentWorkbook(), each Text.StartsWith([Name],"tData205"))[Content],
    SplitSize = 2,
    Join = Table.Join(Sources{1}, "Item", Sources{0}, "Item"),
    ItemCols = List.Transform({1..SplitSize}, each "Item"&Text.From(_)),
    Group = Table.Group(Join, "YesNo", {"All", each let
        _Total = List.Sum(_[Value]),
        _Split = List.Transform(Table.Split(_,SplitSize), each let
            _Sum = List.Sum(_[Value]), 
            _RecI = Record.FromList(_[Item], List.FirstN(ItemCols, List.Count(_[Item])) )
            in Table.FromRecords({[YesNo=_[YesNo]{0}] & _RecI & [Sum=_Sum, #"%age"=_Sum/_Total]}) )
        in Table.Combine(_Split) 
     }),
    CombineAndSort = Table.Sort(Table.Combine(Group[All]), {"YesNo", Order.Descending})
in
    CombineAndSort

