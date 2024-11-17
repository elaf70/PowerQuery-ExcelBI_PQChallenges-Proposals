/*================================
Ref = https://www.linkedin.com/feed/update/

PQ proposal in 3 main steps 
  1) Format date column in Text
  2) Group by [Hall] with unpivot other cols + add in it a 1st Row with the [Hall] description values   
  3) Combine all inner-tables of previous Group result
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData236"]}[Content],
    FormatDate = Table.TransformColumns(Source, {"Date", each Date.ToText(Date.From(_), "d", "Fr-fr")}),
    Group = Table.Group(FormatDate, "Hall", {"G", each let
        _Unpivot = Table.UnpivotOtherColumns(Table.RemoveColumns(_,"Hall"), {}, "Data1", "Data2")//,
        in Table.InsertRows(_Unpivot, 0, {[Data1="Hall", Data2=[Hall]{0}]})
        }),
    Result = Table.Combine(Group[G])
in
    Result
