/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7065899005959901184-Srl1
PQ proposal in 2 main steps
    (1) Group using List.Alternate for splitting items to sum as positives/negatives. 
    (2) Sort
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData080"]}[Content],
    Group = Table.Group(Source, {"Animal"}, {
        {"Sum", each List.Sum([Amount])},
        {"Alternating Sum", each let 
            _L1 = List.Alternate([Amount],1,1,1),
            _L2 = List.Alternate([Amount],1,1,0)
            in List.Sum(_L1)- List.Sum(_L2)} }),
    Sort = Table.Sort(Group,{{"Animal", Order.Ascending}})
in
    Sort


