/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7159395081773301760-PXyn

PQ proposal in 2 main steps
    1) Group by [Pilot) + List.Accumulate to calculate Flight Time & Rest
    2) Expand previous calculation results
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData153"]}[Content],
    Group = Table.Group(Source, {"Pilot"}, {"All" , each
        [TF = List.Accumulate(Table.ToRows(_), 0, (s,c)=>s+Duration.TotalHours(c{2} - c{1})),
        TR = if Table.RowCount(_)=1 then null
            else Duration.TotalHours(List.Max([Flight End])-List.Min([Flight Start])) - TF]
        }),
    Expand = Table.ExpandRecordColumn(Group, "All", {"TF", "TR"}, {"Time Fly", "Time Rest"})
in
    Expand

