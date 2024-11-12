/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7146713102334558208-NF_C

PQ proposal in 3 main steps
    1) Group by {[Emp], [Value]} to select list of {first and last duplicate} or an empty list if there's less than 2 items in each group items 
    2) Back to table from combining distinct Records 
    3) Sort by [Index] to restore original order
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData143"]}[Content],
    Group = Table.Group(Source, {"Emp", "Value"}, {"All", (t)=>let
       _N =  Table.RowCount(t)
       in if (_N>1) then {t{1}, t{_N-1}} else {} }),  
    ToTable = Table.FromRecords(List.Distinct(List.Combine(Group[All]))),
    Sort = Table.Sort(ToTable,{{"Index", Order.Ascending}})
in
    Sort


