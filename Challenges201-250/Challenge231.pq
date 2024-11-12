/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7258326975226789888-mOey

PQ proposal in 5 main steps :...too long😒 - there should be sth simplier 
    1) Buffer T2-Prices in cols
    2) Transform T1-rows in tables with Split,Group & Pivot + add total 
    3) Combine tables+replace Null 
    4) Reorder cols
    5) Add GranTotal row
==================================*/
let
  fxSource = (n)=>Excel.CurrentWorkbook(){[Name=n]}[Content],
  Prices = List.Buffer(Table.ToColumns(fxSource("tData231b"))),
  Source = Table.Sort(fxSource("tData231a"), "Person"),
  Transform = Table.TransformRows(Source, each let 
      _T = Table.FromRows(List.Zip({Text.Split([Items], ", "),List.Transform(Text.Split([Quantity], ", "), Number.From)}), {"I","Q"}),
      _Group = Table.Group(_T, "I", {"V", each List.Sum([Q]) * Prices{1}{List.PositionOf(Prices{0}, [I]{0})} })
      in  Table.FromRows({{[Person]}&_Group[V]&{List.Sum(_Group[V])}}, {"Name"}&_Group[I]&{"Total"}) ),
  Combine = Table.TransformColumns(Table.Combine(Transform), {}, each _??0),
  UsedItems = List.Sort(List.RemoveItems(Table.ColumnName(Combine), {"Name","Total"})),
  OrderCols = Table.ReorderColumns(Combine, {"Name"} & UsedItems & {"Total"}),
  AddTotalR = Table.FromColumns(
      {OrderCols[Name] & {"Total"}} & List.Transform(List.Skip(Table.ToColumns(OrderCols)), each _ & {List.Sum(_)}),
      Table.ColumnNames(OrderCols) )
in
    AddTotalR

