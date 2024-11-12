/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7108662310734315520-wyEd

PQ proposal in 2 main steps
    1. Transform each row by Spliting values in subsets + add Total in each
    2. Convert back to table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData113"]}[Content],
    Transform = Table.TransformRows(Source, each let
        _L = Record.ToList(_),
        _R = List.Accumulate(List.Split(List.Skip(_L), 4), {}, 
            (s,c)=>s & { {_L{0}} & c & {List.Sum(List.LastN(c,2))} } )
        in _R),
    ToTable = Table.FromRows(List.Combine(Transform), 
        List.FirstN(Table.ColumnNames(Source),5)&{"Total"} )
in
    ToTable

