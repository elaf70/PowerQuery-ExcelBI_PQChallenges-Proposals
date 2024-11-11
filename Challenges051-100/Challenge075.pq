/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7060463188269555712-bCnW
PQ proposal in 5 main steps
    1) ...
    2) ...
    3) ... 
    4) ...
    5) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData075"]}[Content],
    Add_IsNull = Table.AddColumn(Source, "IsNull", each [Quantity]=null),
    Group = Table.Group(Add_IsNull, {"IsNull"}, {"Values", each __fxGroupSum(_)}, GroupKind.Local),
      __fxGroupSum = (t)=>Table.Group(t, {"IsNull"}, 
          { {"Quantity", each List.Sum([Quantity]), type number}, 
            {"Amount", each List.Sum([Amount]), type number} 
          }), 
    FilterRows = Table.RemoveColumns(Table.SelectRows(Group, each ([IsNull] = false)), {"IsNull"}),
    Add_GroupName = Table.FromColumns({__GrpNames} & Table.ToColumns(FilterRows), {"Group", "Values"}),
      __GrpNames = List.Transform({1..Table.RowCount(FilterRows)}, each "Group"&Text.From(_) ), 
    Expand = Table.ExpandTableColumn(Add_GroupName, "Values", {"Quantity", "Amount"})
in 
  Expand


