/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7266299517019279361-0C8S

PQ proposal in 1 main step
1) GroupBy [Item] and for each 
 GroupBy [Status] & Count + Sort it to combine inner rows as expected 
==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData238"]}[Content],
  Group = Table.Group(Source, "Item", {"Status", each let 
    _Group = Table.Group(_, "Status", {"Count", each Text.From(Table.RowCount(_))}),
    _Sort = Table.Sort(_Group,{"Status",Order.Descending}),
    _Sep = if Table.RowCount(_Sort)=1 then "-All " else "-"
    in Text.Combine(Table.TransformRows(_Sort, each Text.Combine(Record.ToList(_), _Sep)), ", ") })
in
  Group
