/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7131491217397219328-qa1H

PQ proposal in 4 main steps
    1) Rename Group Column
    2) GroupRow with a Record from all {field,value} where field<>null 
    3) Get the list of all FieldNames for all Groups
    4) Expand Record column
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData131"]}[Content],
    Rename = Table.RenameColumns(Source, {"Column1", "Group"}),
    Group = Table.Group(Rename, {"Group"}, {"All", each let 
      _L = List.Split(List.Combine(List.Skip(Table.ToColumns(_))),2),
      _RN = List.Combine(List.Select(_L, each _{0}<>null))
      in Record.FromList(List.Alternate(_RN,1,1), List.Alternate(_RN,1,1,1)) }),
  FName = List.Union(List.Transform(Group[All], Record.FieldNames)),  
  Expand = Table.ExpandRecordColumn(Group, "All", List.Sort(FName))
in
  Expand











