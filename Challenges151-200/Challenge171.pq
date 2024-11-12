/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7182225510289051649-pQ-K

PQ proposal in 2 main steps
    1) Transform each Lines in List of alternated pairs removing all null ones 
    2) Combine all to Table result
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData171"]}[Content],
    Transform = List.Transform(Table.ToRows(Source), 
         each List.RemoveItems(List.Zip(List.Split(_,3)), {{null,null}})),
    ToTable = Table.FromRows(List.Combine(Transform), {"Col1", "Col2"})
in
    ToTable


