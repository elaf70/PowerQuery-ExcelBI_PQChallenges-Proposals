/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7053215431163092992-ZyKk
PQ proposal in 3 main steps
    1) Group
    2) Combine
    3) Sort
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData070"]}[Content],
    Group   = Table.Group (Source, "HT",
      {"T", each let
        _N = Table.RowCount(_)
        in Table.AddColumn(_, Text.From(_N), each _N) 
      }, 
      GroupKind.Local ),
    Combine = Table.Combine(Table.SelectRows(Group, each [HT]="H")[T]),
    SortCols = Table.SelectColumns(Combine, 
      {"Alpha"} & List.Sort(List.Skip(Table.ColumnNames(Combine),2)))
in
    SortCols


