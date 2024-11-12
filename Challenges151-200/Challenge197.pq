/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7215203817972260864-fmSP

PQ proposal in 4 main steps
    1) Add a new temporary [Index] column to keep original rows order
    2) Group using List.Generate to create new records with Start+End Stock
    3) Combine all these new Records to result table 
    4) Sort it back using temporary column [Index] + remove it 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData197"]}[Content],
    Add_Index = Table.AddIndexColumn(Source, "Index", 0, 1, Int64.Type),
    Group = Table.Group(Add_Index, {"Item","Store"}, 
        {"G", each List.Generate( 
            ()=> [i=0, r=_{i}, ss=r[Stock IN], es=ss-r[Stock OUT]],
            (x)=> x[i]<Table.RowCount(_),
            (x)=> [i=x[i]+1, r=_{i}, ss=x[es], es=ss+r[Stock IN]+r[Stock OUT]],
            each [r] & [#"Start Stock"=[ss], #"End Stock"=[es]])
        }),
    Combine = Table.FromRecords(List.Combine(Group[G])),
    SortBack = Table.RemoveColumns(Table.Sort(Combine, "Index"), "Index")
in
    SortBack


