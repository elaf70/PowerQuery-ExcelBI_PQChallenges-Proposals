/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7075683478507257856-iaP6
PQ proposal in 4 main steps
    (1) Change Age in tens 
    (2) Group 
    (3) Pivot 
    (4) Rename AgeGroup as required 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData087"]}[Content],
    Transform_Age = Table.TransformColumns(Source,
        {"Age", each let 
            _N = Number.IntegerDivide(_,10)
            in if (_N<2) then 1 else if (_N>=8) then 8 else _N, Int64.Type} ),
    Group = Table.Group(Transform_Age, {"Age", "Gender"}, {"Count", each Table.RowCount(_)}),
    Pivot = Table.Pivot(Group, List.Distinct(Group[Gender]), "Gender", "Count", List.Sum),
    Set_AgeGrpName = Table.TransformColumns(Pivot,
        {"Age", each if (_=1) then "<20" else if (_=8) then ">=80" else Text.From(_*10)&"-"&Text.From(_*10+9)})
in
    Set_AgeGrpName


