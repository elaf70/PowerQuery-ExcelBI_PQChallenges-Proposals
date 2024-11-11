/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7015164704989822976-5PP0
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData040"]}[Content],
  IDColNames = List.Select(Table.ColumnNames(Source), each Text.StartsWith(_, "ID")),
  GroupBy = Table.Group(Source, {"Date"}, 
    {"Values", each 
      let
        Step1=Table.ToColumns(Table.SelectColumns(_, IDColNames)),
        Step2=List.Transform(Step1, each {Text.Combine(_,", ")}),
        Step3=Table.FromColumns(Step2, IDColNames)
      in Step3
    }),
    Expand = Table.ExpandTableColumn(GroupBy, "Values", IDColNames, IDColNames),
    ChangeType = Table.TransformColumnTypes(Expand,{{"Date", type date}})
in
  ChangeType

