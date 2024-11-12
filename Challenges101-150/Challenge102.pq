/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7093802873863512064-GwX4

PQ proposal in 6 main steps
    1) ...
    ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData102"]}[Content],

    // Get All Combinaision from a List : 
    //      For 1 list of N items =>  there's (2^N - 1) combinaisons 
    //      Get get each of these number in base2 representation 00000, 00001, 00010, ...  11111
    //          where Each 1 indicate its equivalent in the list    
    fxListSubsets = (L as list) => let
        N = List.Count(L),
        Subsets = List.Transform( {1..Number.Power(2, N)-1},
            (i) => List.Transform({0..N-1},
                (j) => if Number.Mod(Number.IntegerDivide(i, Number.Power(2, j)), 2) = 1
                        then L{j}
                        else null 
                )
            ),
        Concatenate = List.Transform(Subsets, each Text.Combine(List.RemoveNulls(_), ", "))
        in
            Concatenate,
    
    Transform = Table.FromColumns(
        {fxListSubsets(Source[Dept]), fxListSubsets(Source[Name])},
        {"Dept", "Name"} ),
    RemoveDupl = Table.Distinct(Transform, {"Dept"}),
    
    Add_D1C = Table.AddColumn(RemoveDupl, "C1", each Text.Start([Dept], 1)),
    Add_DLen = Table.AddColumn(Add_D1C, "L", each Text.Length([Dept])),
    Sort = Table.Sort(Add_DLen,{{"C1", Order.Ascending}, {"L", Order.Ascending}, {"Dept", Order.Ascending}}),
    SelCols = Table.SelectColumns(Sort,{"Dept", "Name"})
in
    SelCols

