/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7116633309752811520-0aqu

PQ proposal in 4 main steps
    1. Transform each row with List.Accumulate
    2. Back to table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData120"]}[Content],
    Transform = List.Transform(Table.ToRows(Source), each 
        List.Accumulate(List.Skip(_), [r={_{0}}, nn={}, prev=null],
            (s,c)=>let 
                _AvgNN=(c??Number.Round(List.Average(s[nn])))??s[prev]
                in if (c is null) 
                    then [r=s[r]&{_AvgNN??s[prev]}, nn={}, prev=_AvgNN] 
                    else [r=s[r]&{c}, nn=s[nn]&{c}, prev=s[prev]] )[r] 
        ),
    ToTable = Table.FromRows(Transform, Table.ColumnNames(Source))
in
    ToTable
