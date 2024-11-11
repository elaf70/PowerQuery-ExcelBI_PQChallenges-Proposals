/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6977110754982260736-Z_Xz
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData010"]}[Content],
    ToColumns = Table.ToColumns(Source),
    ZipAlternates = List.Zip({
        List.Combine(List.Alternate(ToColumns,1,1,1)), 
        List.Combine(List.Alternate(ToColumns,1,1,0))}),
    ToRecord = Table.FromRecords(
        List.Transform(ZipAlternates, (v)=>Record.FromList(v, {"Column1", "Column2"}))),
    Filter = Table.SelectRows(ToRecord, each ([Column1] <> null))
in
    Filter
