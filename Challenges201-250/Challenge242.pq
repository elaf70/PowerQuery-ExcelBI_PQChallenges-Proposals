/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7271372932021436416-riX5

PQ proposal in 2 main steps
1) Group by [Hall] and Transform each group in List of {Attribute, Value}
2) Combine all of these in the final result table 

==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData242"]}[Content],
  Group = Table.Group(Source, "Hall", 
    {"G", (t)=>let 
      _Unpivot = Table.UnpivotOtherColumns(Table.RemoveColumns(t,{"Hall"}), {}, "A", "V")
      in {{"Hall", t[Hall]{0}}} & Table.ToRows(_Unpivot)
    } ),
  Result = Table.FromRows(List.Combine(Group[G]))
in 
  Result
