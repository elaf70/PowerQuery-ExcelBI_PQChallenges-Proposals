/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-6999943913604046848-83mD
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData028"]}[Content],
    // Group By Greek-Letter + Nb Items of each Group
    Group = Table.Buffer(Table.Group(Source, {"Greek"}, 
        {   {"Data", each _}, 
            {"Count", each Table.RowCount(_)}
        })),
    // Get Max of Items 
    MaxItems = List.Max(Group[Items]),
    // Build List of Records by adding to each Group ones 
    // some extra records with (Null, CurrentLetter) to reach MaxItems in each 
    Transform = List.Transform(Table.ToRecords(Group[[Data],[Count]]),  each
        Table.ToRecords([Data]) 
        & List.Repeat({[Seq=null, Greek=[Data][Greek]{0}]}, MaxItems-[Count]) ),
    // Convert to Table
    ToTable = Table.FromRecords(List.Combine(Transform))
in
    ToTable

