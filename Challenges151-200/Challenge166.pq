/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7174977760031895552-_EFA

PQ proposal in 2 main steps
    1) Group
    2) Sort
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData166"]}[Content],
    CN = {"Company"} & Table.ColumnNames(Source),
    Group = Table.Group(Source, "Tracking No", {"G", 
        each #table(CN, {{[Tracking No]{0}, [Tracking No]{1}, List.Count(List.RemoveNulls([Items])), List.Sum([Amount])}})}, 
        0, (p,c)=>Number.From(c is text)),
    Result = Table.Sort(Table.Combine(Group[G]), "Company")
in
    Result


