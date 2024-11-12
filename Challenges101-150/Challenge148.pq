/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7152147318102728705-6jF2

PQ proposal in 4 main steps
    Trying to avoid as always, Pivot/Unpivot calls if it's to apply on global table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData148"]}[Content],
    Split = Table.TransformColumns(Source, {"Fruits", each Text.Split(_, ", ")}),
    Expand = Table.ExpandListColumn(Split, "Fruits"),
    Sort = Table.Sort(Expand,{"Fruits", Order.Ascending}),
    Group = Table.Group(Sort, {"Fruits"}, {"Count", each Record.FromList({Table.RowCount(_)}, {_[Fruits]{0}})}),
    Expand2 = Table.ExpandRecordColumn(Group, "Count", Group[Fruits])
in
    Expand2


