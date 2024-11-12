/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7240937984114491393-UPyv

PQ proposal in 1 main steps 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData218"]}[Content],
    Group = Table.Group(Source, "Project", {
        {"Completed Tasks"      , each Text.Combine(Table.SelectRows(_, each [Completion Date]<>null)[Tasks], ", ")},
        {"Non Complete Tasks"   , each Text.Combine(Table.SelectRows(_, each [Completion Date]= null)[Tasks], ", ")} })
in
    Group


