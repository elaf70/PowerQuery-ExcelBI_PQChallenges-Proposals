/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6979288688568467456-ZPGj
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
   #"▶ColumnNames" = Table.ColumnNames(Source),
   #"▶Subj_Columns" = List.Select(#"▶ColumnNames", each Text.StartsWith(_,"Subject")),
   #"▶Mark_Columns" = List.Select(#"▶ColumnNames", each Text.StartsWith(_,"Marks")),

   Source = Excel.CurrentWorkbook(){[Name="tData11"]}[Content],
   Add_Merge = Table.AddColumn(Source, "Merge", each let
      _Subj = Record.ToList(Record.SelectFields(_, #"▶Subj_Columns")),
      _Mark = Record.ToList(Record.SelectFields(_, #"▶Mark_Columns")),
      _Table = Table.FromColumns({_Subj, _Mark}, {"Subject", "Mark"})
      in Table.SelectRows(_Table, each ([Subject] <> null))
   ),
   Expand = Table.ExpandTableColumn(Add_Merge[[ID],[Merge]], "Merge", {"Subject", "Mark"}),
   ChangeType = Table.TransformColumnTypes(Expand,{{"Mark", Int64.Type}}),
   Pivot = Table.Pivot(ChangeType, List.Distinct(ChangeType[Subject]), "Subject", "Mark")
in
   Pivot
