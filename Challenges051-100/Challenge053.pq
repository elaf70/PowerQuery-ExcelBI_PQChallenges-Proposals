/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7032559322454736896-OcwA
PQ proposal in 2 main steps
1) Group by {"Country", "State", "City"} of T1 with a dynamic creation of all new columns function built from T2
2) Sort
==================================*/
let
    Source1 = Excel.CurrentWorkbook(){[Name="tData053_1"]}[Content],
    Source2 = Excel.CurrentWorkbook(){[Name="tData053_2"]}[Content],
    T2_ListAllItems = List.RemoveNulls(List.Combine(Table.ToColumns(Source2))),
    T2_ColNames = Table.ColumnNames(Source2),
    Group = Table.Group(Source1, {"Country", "State", "City"},
       List.Transform (T2_ColNames,
          each {_, (GrpSetT1)=>let
            _ColNames = List.RemoveNulls(Table.Column(Source2, _)),
            _Sum = List.Accumulate(_ColNames, 0, 
              (s,c)=>let 
                _SumCol = List.Sum(Table.Column(GrpSetT1,c)),
                _CountItem =  List.Count(List.Select(T2_ListAllItems, each _=c))
                in s + _SumCol/ _CountItem )
            in _Sum, type number })
    ),
    Sort = Table.Sort(Group,{{"Country", Order.Ascending}, {"State", Order.Ascending}, {"City", Order.Ascending}})
in
    Sort