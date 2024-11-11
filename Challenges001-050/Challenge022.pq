/*================================
Ref = https://www.linkedin.com/posts/excelbi_excel-excelchallenge-powerquerychallenge-activity-6992334759410446336-biTA
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData022_1"]}[Content],
    Data = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData022_2"]}[Content]),
    Transform = List.Accumulate(
       List.Sort(List.Distinct(Data[Date])), 
       Table.Sort(Source, "Employees"), 
       (s,c)=> Table.AddColumn(s, c, (e)=>let
           _RelatedDates = Table.SelectRows(Data, (d)=>(d[Date]=c and d[Employees]=e[Employees]))
           in Table.RowCount(_RelatedDates))
       )
in
    Transform

