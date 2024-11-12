/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7177514489527877632-m3FP

PQ proposal in 2 main steps
    1) Group by Store + 2 List calls Sort + Acculumate for concatenation 
    2) Combine inner tables 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData168"]}[Content],
    Group = Table.Group(Source, "Store", {"G", each let
        S = List.Sort([Item]), 
        T = List.Accumulate(List.Skip(S), {S{0}}, (s,c)=>s & {List.Last(s)&"/"&c})
        in Table.FromColumns({[Store], T}, {"Store", "Item"}) 
        } ),
    Result = Table.Combine(Group[G])
in
    Result


