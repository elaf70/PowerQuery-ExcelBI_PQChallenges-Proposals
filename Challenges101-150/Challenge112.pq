/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7106486447506898944-k-br

PQ proposal in 2 main steps
    1) GroupBy Task with 1 inner GroupBy Role 
    to combine relative Records with ListofPerson & their sum of Hours
    2) Expand those Records
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData112"]}[Content],
    Group1 = Table.Group(Source, {"Task ID", "Role"}, 
        {"R", each Record.FromList(
            {Text.Combine([Person], ", "), List.Sum([Hours])},
            {[Role]{0}&"s", [Role]{0}& " Hours"})
        }),
    Group2 = Table.Group(Group1, {"Task ID"}, {"R", each Record.Combine([R])}),
    Expand = Table.ExpandRecordColumn(Group2, "R", 
        List.Union(List.Transform(Group2[R], each Record.FieldNames(_)))) 
in
    Expand


