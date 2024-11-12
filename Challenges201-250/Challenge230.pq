/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7256152635278323712-rPWt

PQ proposal in 2 main steps
    1) Split by subset of 2 cols to Merge all of them into 1 single table 
    2) Group by [Month] : local.kind + Sum([sale]) of each
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData230"]}[Content],
    Transform = Table.FromRows(List.Combine(List.Transform(List.Split(Table.ToColumns(Source),2), List.Zip)), {"Month", "Sale"}),
    Group = Table.Group(Transform, "Month", 
        {"Sale", each List.Sum([Sale])}, GroupKind.Local, (p,c)=>Number.From(not(Text.StartsWith(c,"Week"))))
in
    Group


