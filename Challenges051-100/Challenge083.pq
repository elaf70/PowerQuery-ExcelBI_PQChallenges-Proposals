/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7070610050200813568-TyRA
PQ proposal in 3 main steps
    1) Transform columns = PrefixQuestionName + Merge Answers(Y/N)
    2) Group by Student with Pivot on Question + Expand result, 
    3) Add Score
 ==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData083"]}[Content],
    Transform = let 
        _tr = Table.TransformRows(Source, each
            [ Student=[Student], Q = "Q" & Text.From([Question No]),
            V = if ([Selected Answer]=[Right Answer]) then "Y" else "N" ]) 
            in Table.FromRecords(_tr),
    Group =Table.Group(Transform, "Student", {"Data", each Table.Pivot(_, List.Distinct(_[Q]),"Q", "V") } ),
    Expand= Table.ExpandTableColumn(Group, "Data", List.Distinct(Transform[Q]) ),
    Add_Score = Table.AddColumn(Expand, "% Score", each let
        _lv = List.Skip(Record.FieldValues(_)),
        _Score = List.Count(List.FindText(_lv, "Y")) / List.NonNullCount(_lv)
        in Number.Round(_Score,2) ) 
in
    Add_Score


