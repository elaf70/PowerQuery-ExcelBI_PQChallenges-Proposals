/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7103587351402049536-hgu7

PQ proposal in 4 main steps
    1) Sort + AddIndex
    2) Convert to Records + Buffered the list
    3) Select only record matching expected Criterias with Relative Record 
    = Previous/Next one depending on Category - with index + try 
    4) Back To table from selected records 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData109"]}[Content],
    Sort = Table.Sort(Table.TransformColumnTypes(Source,{{"Emp ID", Int64.Type}, {"Date Time", type datetime}}),
        {{"Emp ID", Order.Ascending}, {"Date Time", Order.Ascending}} ),
    Add_Index = Table.AddIndexColumn(Sort, "Index", 0, 1, Int64.Type),
    Records = List.Buffer(Table.ToRecords(Add_Index)),
    Select = List.Select(Records, each let
        _Relative = try Records{_[Index] + (if _[Category]="Entry" then +1 else -1)} otherwise _
        in _[Emp ID]<>_Relative[Emp ID] or _[Category]=_Relative[Category]
    ),
    ToTable = Table.FromRecords(Select, Table.ColumnNames(Source))
in
    ToTable


