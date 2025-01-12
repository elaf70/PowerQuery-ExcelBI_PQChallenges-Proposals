/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7284056663161004032-_Rg_

PQ proposal in 2 main steps
  1) Group in LocalMode to create inner tables from records arrangements
  2) Combine tables

==================================*/

let
  Source = Excel.CurrentWorkbook(){[Name="tData252"]}[Content],
  Group = Table.Group(Source, "Column1", 
    {"G", each let
      R1 = Record.SelectFields(_{0}, {"Store No", "Store Name"}),
      R = Table.ToRecords(Table.RemoveColumns(Table.PromoteHeaders(Table.RemoveFirstN(_)), {"Column2"}))
      in Table.FromRecords(List.Transform(R, each R1 & _))
    }, 
    GroupKind.Local, (x,y)=>Number.From(y=null)),
  Result = Table.Combine(Group[G])
in
  Result
