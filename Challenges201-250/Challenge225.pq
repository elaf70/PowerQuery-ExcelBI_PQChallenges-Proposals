/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7250716820763013121-J5GG

PQ proposal in 2 main steps
    1) Group by {Group,Dept} and for each create a table FromRows with 
        the append of their values from 1st rec. + all {EmpID,Name} row-values
    2) Combine all these grouping inner-tables
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData225"]}[Content],
    Group = Table.Group(Source, {"Group", "Dept"},           {"G", each  Table.FromColumns({{[Group]{0}}&[Emp ID], {[Dept]{0}}&[Name]}) }),
    Combine = Table.Combine(Group[G])
in
    Combine


