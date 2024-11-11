/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7062999908445753344-SwiX
PQ proposal in 4 main steps
    (1) Transform = split by CharacterTransition into Records
    (2) Expand Records 
    (3) Group by Letter with Sumof Quantity
    (4) Sort 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData077"]}[Content],
    Transform = Table.TransformColumns(Source, {"String", each let 
        _S = Text.Replace(_, " ", ""),
        _Split=Splitter.SplitTextByCharacterTransition({"0".."9"}, {"A".."Z"})(_S)
        in List.Transform(_Split, each 
            Record.FromList( Splitter.SplitTextByCharacterTransition( {"A".."Z"},{"0".."9"})(_), {"Letters", "Value"})) }),
    Expand = Table.ExpandRecordColumn( Table.ExpandListColumn(Transform, "String"), "String", {"Letters", "Value"}),
    ChangeType = Table.TransformColumnTypes(Expand,{{"Value", Int64.Type}}),
    Group = Table.Group(ChangeType, {"Letters"}, {{"Quantity", each List.Sum([Value])}}),
    Sort = Table.Sort(Group,{{"Letters", Order.Ascending}})
in 
    Sort

























