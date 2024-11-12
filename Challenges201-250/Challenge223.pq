/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7248180131523923969-hdUe

PQ proposal in 2 main steps
    1) Group by <Group> to split each related-content into inner-table of 2 pairs {Code,Value}, with List.Zip to reorder with code-items first 
    2) Expand it
==================================*/
let
    CN = {"Code1", "Code2", "Value1", "Value2"},

    Source = Excel.CurrentWorkbook(){[Name="tData223"]}[Content],
    ChgType = Table.TransformColumnTypes(Source,{{"Type", type text}, {"Code", type text}}),
    Group = Table.Group(ChgType, "Group", {"G", each let
        _CombineCode = Table.CombineColumns(_,{"Type", "Code"},Combiner.CombineTextByDelimiter(""),"Code")[[Code],[Value]],
        _Split = List.Transform(List.Split(Table.ToRows(_CombineCode),2), each 
            List.Combine(List.Zip(if List.Count(_)=2 then _ else _ & {{null,null}})))
        in Table.FromRows(_Split, CN) }),
    Expand = Table.ExpandTableColumn(Group, "G", CN)
in
    Expand



