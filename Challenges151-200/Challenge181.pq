/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7194914135711637504-ePRl

PQ proposal in 2 main steps + 1 defined function
    1) Group by [Column1] and call fxFunction to transform it an inner table with Unpivot
    2) Combine
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData181"]}[Content],
    fxTransformInnerTable = (t as table) => let
       R0 = Record.ToList(t{0}),
       RenameC = Table.RenameColumns(t, List.Zip({Table.ColumnNames(t), {"Name"} & List.Skip(R0)})),
       Unpivot = Table.UnpivotOtherColumns(Table.RemoveFirstN(RenameC,1), {"Name"}, "Data", "Value"),  
       AddC = Table.AddColumn(Unpivot, "Date", each Date.From(R0{0}), type date)
       in AddC, 
    Group = Table.Group(Source, "Column1", {"G", each fxTransformInnerTable(_)}, 
        GroupKind.Local, (x,y)=>try Number.From((Date.From(y) is date)) otherwise 0),
    Combine = Table.Combine(Group[G], {"Date", "Name", "Data", "Value"})
in
    Combine


