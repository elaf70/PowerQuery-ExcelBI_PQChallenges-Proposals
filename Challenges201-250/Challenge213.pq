/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7235496423708889088-UcY-

PQ proposal in 5 main steps
    (1) Merge 2 tables, 
    (2)Tranform columns {Group,Item} and Expand
    (3) Pivot [Item] 
    (4+5) Add Total Row & Col
==================================*/
let
    fxGetTable = (t)=>Excel.CurrentWorkbook(){[Name=t]}[Content],
    Merge = fxGetTable("tData213a") & fxGetTable("tData213b"), 
    TransformCols = List.Accumulate({"Group", "Item"}, Merge, 
        (s,c)=>Table.ExpandListColumn(Table.TransformColumns(s, {c, each Text.Split(_, ", ")}), c)),
    Pivot = Table.Pivot(TransformCols, List.Distinct(TransformCols[Item]), "Item", "Stock", List.Sum),
    CN = Table.ColumnNames(Pivot) & {"Total"},
    Add_TotalCol = let
        _Values = List.Transform(Table.ToRows(Pivot), each _ & {List.Sum(List.Skip(_))})
        in Table.FromRows(_Values, CN),
    Add_TotalRow = let 
        _Values = List.Transform(List.Skip(Table.ToColumns(Add_TotalCol)),List.Sum)
        in Table.Combine({Add_TotalCol, Table.FromRows({{"Total"} & _Values}, CN)})
in
    Add_TotalRow


