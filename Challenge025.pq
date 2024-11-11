/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-6997046004357193728-fzr-
PQ proposal in 2 main steps
1) Group by [ID] to create for each 1 inner table with required items 
2) Expand it
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData025"]}[Content],
    Group = Table.Group(Source, {"ID"}, {"Data", each let 
        _IsFirstZTemp  = try (Table.SelectRows(_, each [Event]="Z")[Category]{0}="Temporary") otherwise false
        in Table.AddColumn(_, "Value", each if (_IsFirstZTemp) then "Yes" else "No") }),
    Expand = Table.ExpandTableColumn(Group, "Data", {"Event", "Category", "Value"})
in
    Expand

