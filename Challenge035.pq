/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7009728886896771072-fEIw
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData035"]}[Content],
    Transform = List.Accumulate(Table.ToRows(Source), {}, (s,c)=>s&{c}&List.Repeat({{null,null}}, c{1})),
    ToTable = Table.FromRows(Transform, Table.ColumnNames(Source))
in
    ToTable

