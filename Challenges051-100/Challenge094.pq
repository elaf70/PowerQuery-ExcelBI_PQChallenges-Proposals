/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7083656011600846848-ZQCt
PQ proposal in 3 main steps
    1) ...
    2) ...
    3) ... 
==================================*/
let
    Source = Table.TransformColumnTypes(
        Excel.CurrentWorkbook(){[Name="tData094_2"]}[Content],
        {"Date", type date}),
    TData  = Table.Buffer(Table.TransformColumnTypes(
        Excel.CurrentWorkbook(){[Name="tData094_1"]}[Content],
        {"Use Date", type date})),
    Add_Transform = Table.AddColumn(Source, "T", each let
        _R = _,
        _TB = Table.SelectRows(TData, each ([Machine]=_R[Machine] and [Use Date]<=_R[Date]) ),
        _B = List.Average(Table.LastN(_TB,4)[Use Mins]),
        _TA = Table.SelectRows(TData, each ([Machine]=_R[Machine] and [Use Date]>_R[Date]) ),
        _A = List.Average(Table.FirstN(_TA,4)[Use Mins])
        in [B=Number.Round(_B,0), A=Number.Round(_A,0)]
        ),
    Expand = Table.ExpandRecordColumn(Add_Transform, "T", {"B", "A"},        {"Avg of 4 Use Mins Before or On", "Avg of 4 Use Mins After"})    
in
    Expand


