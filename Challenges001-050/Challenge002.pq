/*===============================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6968427924391649280-WZey
PQ proposal in 2 main-steps
1) Transform rows by Combining all their values + use List.Accumulate to create 2 sublists of Digits and NonDigits items 
2) Create result table with it 
=================================*/

let
    Source = Excel.CurrentWorkbook(){[Name="tData002"]}[Content],
    Transform = List.Accumulate(List.Combine(Table.ToRows(Source)), {{},{}}, 
      (s,c)=>if ((try Number.From(c) otherwise null)=null) 
            then {s{0}&{c}, s{1}}
            else {s{0}, s{1}&{c}} ),
    ToTable = Table.FromColumns(Transform,{"Alpha", "Number"})
in
    ToTable
