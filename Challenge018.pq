/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6987259798832459776-0NZi
PQ proposal in 4 main steps
1) Create List of Quarters for eachRow with 1 List.Generate sequence
2) Filter those with at least 2 Quarters 
3) Expand + GroupBy with a count 
4) Finally Sort to get final result
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData018"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"In Quarter", type text}, {"Out Quarter", type text}}),
    Transform = Table.AddColumn(ChangeType, "Quarter", 
        each let 
            fxRecordFromText = (s as text) => 
                let in [y=Number.From(Text.Middle(s,1,4)),q=Number.From(Text.Middle(s,6,1))],
            _Out = fxRecordFromText([Out Quarter]),
            _Result = List.Generate(
                ()=>fxRecordFromText([In Quarter]),
                each (_[y]<_Out[y] or (_[y]=_Out[y] and _[q]<_Out[q])),
                each if (_[q]=4) then [y=[y]+1, q=1] else [y=[y],q=[q]+1],
                each "Y" & Text.From(_[y]) & "Q" & Text.From(_[q])
            )
        in _Result),
    Filter = Table.SelectRows(Transform, each (List.Count([Quarter])>0)),
    Expand = Table.ExpandListColumn(Filter[[Quarter]], "Quarter"),
    Group = Table.Group(Expand, {"Quarter"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    Sort = Table.Sort(Group,{{"Quarter", Order.Ascending}})
in
    Sort

