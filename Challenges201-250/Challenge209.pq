/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7230423533481336833-udfN

PQ proposal in 6 main steps ...

==================================*/
let
    Sources = Table.SelectRows(Excel.CurrentWorkbook(), each Text.StartsWith([Name],"tData209"))[Content],
    T1_SplitTask = Table.TransformColumns(Sources{0}, {"Task", each Text.Split(_,", ")}),
    T1_ExpandTask = Table.ExpandListColumn(T1_SplitTask, "Task"),
    JoinT2 = Table.Join(T1_ExpandTask, "Task", Sources{1}, "Task"),
    Add_End = Table.AddColumn(JoinT2, "End", each Date.AddDays([Start Date], [Duration Days])),
    Group = Table.Group(Add_End[[Process],[Owner],[End]], {"Process","Owner"}, {"End", each List.Max([End])}),
    Pivot = Table.Pivot(Group, List.Sort(List.Distinct(Group[Owner])), "Owner", "End", each DateTime.ToText(List.Max(_), "yyyy-MM-dd"))
in
    Pivot



