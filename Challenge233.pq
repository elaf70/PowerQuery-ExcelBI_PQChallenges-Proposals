/*===============================
Ref = https://www.linkedin.com/feed/update/urn:li:activity:7260888179526795264
PQ proposal in 4 main steps
1) Transform by rows to split values + add Status & Amount calcs
2) Pivot [Fruit] into columns
3) Sort : row by [Status] and columns by Fruit-names
4) Add the 2 last column and Row for GranTotal 
=================================*/

let
  Source = Excel.CurrentWorkbook(){[Name="tData233"]}[Content],
  Transform = Table.FromRows(
    List.TransformMany(Table.ToRows(Source), 
      each List.Split(List.Skip(_), 3), 
      (o,_)=>{o{0}, if (_{0}="Y") then "Shipped Amount" else "Not Shipped Amount",_{1}*_{2}}),
    {"F","Status","Amt"} ),
  Pivot = Table.Pivot(Transform, List.Distinct(Transform[F]), "F", "Amt", each List.Sum(_)??0),
  Sort = Table.ReorderColumns(Table.Sort(Pivot,{{"Status", Order.Descending}}), 
    {"Status"} & List.Sort(List.Skip(Table.ColumnNames(Pivot))) ),
  AddGTRowCol = let
    _AddGTCol = Table.FromRecords(Table.TransformRows(Sort, each _ & [Total=List.Sum(List.Skip(Record.ToList(_)))])),
    _AddGTCNum = List.Transform(List.Skip(Table.ToColumns(_AddGTCol)), each _ & {List.Sum(_)})
    in Table.FromColumns({_AddGTCol[Status]&{"Total"}} & _AddGTCNum, Table.ColumnNames(_AddGTCol))
in
  AddGTRowCol
