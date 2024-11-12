/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7172441045060698112-_fUA

PQ proposal in 2 main steps
    1) Group 
    2) Expand 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData164"]}[Content],
    Group = Table.Group(Source, "Group", {{"All", each let
        S = List.Split(List.Skip(Table.ToColumns(_)), 2),
        T = List.Transform(List.Zip(S), each let
            T = Table.FromColumns(_, {"Number", "Value"})
            in Table.SplitColumn(T, "Number", 
                Splitter.SplitTextByPositions({0, 2}, false), {"Type", "Code"}) )
        in Table.Combine(T) }}),
    Expand = Table.ExpandTableColumn(Group, "All", {"Type", "Code", "Value"})
in
    Expand


