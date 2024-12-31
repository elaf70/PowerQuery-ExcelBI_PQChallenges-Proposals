/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7271010533556826112-plrm

PQ proposal in 3 main steps
1) Get ColumnNames
2) Transform each row with Split text + Table.FromColumns & FillDown
3) Combine all of these as final result

==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData241"]}[Content],
  CN = Table.ColumnNames(Source),
  Transform = List.Transform(Table.ToRows(Source), each let
    _Split = List.Transform(_, each try Text.Split(_, ", ") otherwise {})
    in Table.FillDown(Table.FromColumns(_Split, CN), {"Group"}) ), 
  Result = Table.Combine(Transform)
in
  Result 
