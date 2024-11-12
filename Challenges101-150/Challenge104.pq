/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7096339592924082176-GkeG

PQ proposal in 4 main steps
    1) Group by [Company] and Transform each row with Text.Split 
    2) ToTable
    3) FillDown [Company]+[PO No]
    4) ChangeType
    ==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData104"]}[Content],
    Group = Table.Group(Source, {"Company"}, {"Data", each let
        _C = Table.TransformRows(_, each let
            in List.Zip(List.Transform(Record.ToList(_), each Text.Split(_, " / ")))
            ),
        _Combine = List.Combine(_C),
        _Total = List.Sum(List.Transform(_Combine, each Number.FromText(_{4})))
         
        in _Combine & {{null, "Total Items", List.Count(_Combine), "Total Amount",_Total}} }),
    ToTable = Table.FromRows(List.Combine(Group[Data]), Table.ColumnNames(Source)),
    Fill = Table.FillDown(ToTable,{"Company", "PO_No"}),
    ChangeType = Table.TransformColumnTypes(Fill,{{"PO_Line_No", Int64.Type}, {"PO_Amount", Int64.Type}})
in
    ChangeType


