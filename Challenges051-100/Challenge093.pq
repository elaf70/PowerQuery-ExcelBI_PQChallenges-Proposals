/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7083293619222118400-rUOH
PQ proposal in 2 main steps
    (1) Transform in New Record perRow with Max & Years 
    (2) Expand it
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData093"]}[Content],
    Add_Transform = Table.AddColumn(Source, "R", each let
        _T = Table.FromColumns({List.Skip(Record.ToList(_)), List.Skip(Record.FieldNames(_))}, {"V","F"} ),
        _Max = List.Max(_T[V]),
        _FList = Text.Combine(Table.SelectRows(_T, each [V]=_Max)[F], ", ")
        in [#"Max Amt"=_Max, #"Max Years"=_FList]),
    Expand = Table.ExpandRecordColumn(Add_Transform, "R", {"Max Amt", "Max Years"})
in
    Expand


