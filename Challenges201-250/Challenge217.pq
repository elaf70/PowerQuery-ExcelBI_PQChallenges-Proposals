/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7240571624696012800-4aqc

PQ proposal in 2 main steps
    1) Transform Rows to Get list of Amounts/Month + total for each user
    2) Create Result with Table.FromColumns (~transpose)
        with 1st col = ListMonth from original table ColumnNames + {"Total"} 
 ==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData217"]}[Content],
    TRows = Table.TransformRows(Source, (r)=>let
        L = List.Transform(List.Skip(Record.ToList(r),2), each _ * r[Amt])
        in L & {List.Sum(L)} ),
    Result = Table.FromColumns( 
        {List.Skip(Table.ColumnNames(Source),2) & {"Total"}} & TRows,
        {"Month"} & Source[Customer])
in
    Result



