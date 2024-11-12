/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7233322200735350784-cwDl

PQ proposal in 3 main steps
    1) Transform each rows of T2 by selecting related rows in T1 and calculate Amount of %Commision from Sales
    2) Group by [Name] to calculate Amount for each
    3) Sort by Amount
==================================*/
let
    Sources = Table.SelectRows(Excel.CurrentWorkbook(), each Text.StartsWith([Name],"tData212"))[Content],
    T1 = Table.Buffer(Sources{0}),
    Transform = List.Transform(Table.ToRows(Sources{1}), each let
        _Sales = _{1},
        _Codes = List.RemoveNulls(List.Skip(_,2)),
        _T = Table.SelectRows(T1, each List.Contains(_Codes,[Code]))
        in Table.AddColumn(_T,"Amount", each _Sales*[Commission]) ),
    Group = Table.Group(Table.Combine(Transform), "Name", {"Amount", each List.Sum([Amount])}),
    Sort = Table.Sort(Group,{{"Amount", Order.Descending}})
in 
    Sort


/*
// SOLUTION 2 : Unpivot/Join/GRoup/Sort
let
    Sources = Table.SelectRows(Excel.CurrentWorkbook(), each Text.StartsWith([Name],"tData212"))[Content],
    Unpivot = Table.UnpivotOtherColumns(Sources{1}, {"Deal", "Sales"}, "Attribute", "Value"),
    Join = Table.Join(Unpivot,"Value", Sources{0}, "Code"),
    Group = Table.Group(Join, "Name", {"Amount", each List.Sum(Table.AddColumn(_,"V",each [Sales]*[Commission])[V])}),
    Sort = Table.Sort(Group,{{"Amount", Order.Descending}})
in 
    Sort
*/