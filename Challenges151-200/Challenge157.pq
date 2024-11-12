/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7164466865094828032-wdnp

PQ proposal in 2 main steps
    1) Group & transform inner tables cols except 1st with List.Accumulate
    2) Combine inner tables 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData157"]}[Content],
    CN = Table.ColumnNames(Source),
    NCol= Table.ColumnCount(Source),
    Transform = List.Accumulate(Table.ToRows(Source), [r={}, prev={-1}], (s,c)=>let
        _Values = if (c{0} = s[prev]{0}) 
            then List.Transform(List.Zip({c,s[prev]}), each if(_{0}=_{1}) then null else _{0})
            else List.Repeat({null}, NCol)
        in [r=s[r]&{{c{0}} & List.Skip(_Values)}, prev=c]),
    ToTable = Table.FromRows(Transform[r], CN)
in
    ToTable


