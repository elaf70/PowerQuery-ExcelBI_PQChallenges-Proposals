/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7012265601683337216-vs0P
PQ proposal in 3 main steps
1) ...
2) ...
3) ... 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData037"]}[Content],
    Add_Family = Table.AddColumn(Source, "Family", 
        each List.Reverse(Splitter.SplitTextByDelimiter(" ", QuoteStyle.None)([Name])){0}?, type text),
    Group = Table.Group(Add_Family, {"Family"}, {"Members", 
        each if (List.Count(_[Name])>1) then Text.Combine(_[Name],", ") else null} ),
    Filter = Table.SelectRows(Group, each ([Members] <> null))
in
    Filter

