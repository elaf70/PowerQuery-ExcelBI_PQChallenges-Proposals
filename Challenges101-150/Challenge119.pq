/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7116270919244820480-ewdG

PQ proposal in 3 main steps
    1. Transform with List Accumulate to add GroupIndex to each Row
    2. ToTable FromRows
    3. Transform "Group" column from number to text 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData119"]}[Content],
    AddGroup = List.Accumulate(Table.ToRows(Source), [t=0, g=1, r={}], (s,c)=>let
        _t = s[t] + c{1}
        in if (_t<=20) 
            then [t=_t,   g=s[g],   r=s[r] & {c & {s[g]}}] 
            else [t=c{1}, g=s[g]+1, r=s[r] & {c & {s[g]+1}}] )[r],
    ToTable = Table.FromRows(AddGroup, Table.ColumnNames(Source)&{"Group"}),
    TransformGroup = Table.TransformColumns(ToTable, {"Group", each "Group" & Text.From(_)})
in
    TransformGroup


