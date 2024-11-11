/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7068435721157525504-HdeW
PQ proposal in 3 main steps
     1) Sort 
     2) TransformToList 
     3) Back To Result Table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData082"]}[Content],
    Sort = Table.Sort(Source,{{"Assignee", Order.Ascending}, {"Priority", Order.Ascending}}),
    TransformToList = List.Accumulate(List.Distinct(Table.ToRows(Sort)),{},(s,c)=>s&{c{0},c{1}}),
    ToTable = Table.FromColumns({TransformToList},{"Result"})
in
    ToTable


