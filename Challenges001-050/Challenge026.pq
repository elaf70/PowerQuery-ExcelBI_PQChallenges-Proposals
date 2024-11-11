/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-6997406715126145024-TnyW
PQ proposal in 2 main steps
1) Group by [ProcessID] to create for each a list of required items
2) Expand it in new rows for each 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData026"]}[Content],
    Group = Table.Group(Source, "Process ID", {"Step", each let
        _NbRows = Table.RowCount(_)
        in if (_NbRows=1) then {"Start/End"} 
            else {"Start"} & List.Transform({1..(_NbRows-2)}, each"Step"&Text.From(_)) & {"End"} }),
    Expand = Table.ExpandListColumn(Group, "Step")
in
    Expand

