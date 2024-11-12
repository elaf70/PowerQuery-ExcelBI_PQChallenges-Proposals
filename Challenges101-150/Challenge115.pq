/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7111197494243438592-A7zZ

PQ proposal in 3 main steps
    1) Filter Num items 
    2) "Transpose" each relative name & value to separate row values 
    3) Back to table, adjusting each of these new rows values 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData115"]}[Content],
    FilterNum  = List.Accumulate(List.Split(Table.ToRows(Source),2), {}, 
        (s,c)=>s & {List.Select(List.Zip(c), each _{1} is number)} ),
    ToTable = Table.Combine(List.Transform(FilterNum, each Table.FromColumns(_)))
in
    ToTable


