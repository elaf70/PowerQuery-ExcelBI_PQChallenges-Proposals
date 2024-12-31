/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7268836106178084864-XpKH

PQ proposal in 3 main steps
1) Transform each row on an inner table Text.Split (x2)
2) Generate List of sorted ColumnNames 
3) Expand inner-tables as final result
 
==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData240"]}[Content],
  Transform = Table.TransformColumns(Source, {"Items", each let
    cv = List.Combine(List.Transform(Text.Split(_, ", " ), each Text.Split(_, ": ")))
    in Table.FromRows({List.Alternate(cv, 1, 1)}, List.Alternate(cv, 1, 1,1))      }),
  CNames = List.Sort(List.Distinct(List.Combine(List.Transform(Transform[Items], each Table.ColumnNames(_))))), 
  Expand = Table.ExpandTableColumn(Transform, "Items", CNames)
in
  Expand
