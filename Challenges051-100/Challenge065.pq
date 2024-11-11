/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7047780975623192576-3IWE
PQ proposal in 4 main steps
    1)RemoveDuplicate, 2)Sort, 3)ListAccumulate on each Row Items and then 4) Back ToTable as Result
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData065"]}[Content],
    RemoveDuplicate = Table.Distinct(Source),
    Sort = Table.Sort(RemoveDuplicate,Table.ColumnNames(RemoveDuplicate)),
     //----  Previous way with  List.Accumulate(Table.ToRows(Sort), {}, (s,c)=>s&c),
    Combine = List.Accumulate(Table.ToRows(Sort), {}, (s,c)=>s&c),
    ToTable = Table.FromList(Combine, Splitter.SplitByNothing(), {"Result"})
in
    ToTable


