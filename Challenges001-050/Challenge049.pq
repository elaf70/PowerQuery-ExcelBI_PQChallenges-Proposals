/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7027485892076351488-OJ2b
PQ proposal in 3 main steps
1) ...
2) ...
3) ... 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData049"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"From", type datetime}, {"To", type datetime}}),
    Group = Table.Group(ChangeType, {"State"}, {"Data", each let
        _L = List.Combine(Table.ToRows([[From], [To]])),
        _Transform = List.Accumulate(List.Skip(_L), [r={_L{0}}, p=_L{0}], 
            (s,c)=>[r=s[r] & {Number.Round(Number.From(c-s[p])*24*60,0)}, p=c])[r],
        _FN = {"From"} & List.Transform({1..List.Count(_Transform)-1}, each "Time" & Number.ToText(_))
        in Record.FromList(_Transform, _FN) }),
    Expand = Table.ExpandRecordColumn(Group, "Data", 
         List.Distinct(List.Combine(List.Transform(Group[Data], each Record.FieldNames(_)))) )
in
    Expand

