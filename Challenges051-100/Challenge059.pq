/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7040169473244315649-o02u
PQ proposal in 5 main steps
    1.Restore Seq num
    2.Partition Table 
    3.Transform part to build each Col (header incl.) 
    4.Construct Result from Cols 
    5.Reorder Col
==================================*/

// SOLUTION 1 : Without PIVOT
//============================
let
    Source = Excel.CurrentWorkbook(){[Name="tData059"]}[Content],
    FillUp_Seq = Table.FillUp(Table.ReplaceValue(Source,"#",null,Replacer.ReplaceValue,{"Seq"}),{"Seq"}),

    Part=Table.Partition(FillUp_Seq, "Seq", List.Last(FillUp_Seq[Seq]), each _),
    NbOptMax = List.Max(List.Transform(Part, Table.RowCount))-1,
    Transform = List.Transform(Part, each let
        _q = Text.Combine({Text.From([Seq]{0}), ".", Text.From([Question]{0})}),
        _o = List.Skip(_[Question]),
        _n = List.Repeat({null} , NbOptMax-List.Count(_o)+1),
        _a = Table.SelectRows(_, each ([Correct] = "Y"))[Question]
        in List.Combine({{_q},_o,_n,_a}) ),
    FirstCol = {"0.Options"} & List.Transform({1..NbOptMax}, (o) => "Option " & Text.From(o) ) & {null, "Correct Answer"},
    Result = Table.PromoteHeaders(Table.FromColumns(Transform & {FirstCol})),
    Reorder = Table.ReorderColumns(Result,List.Sort(Table.ColumnNames(Result))),
    Rename = Table.RenameColumns(Reorder,{{"0.Options", "Options"}})
in
    Rename

/*
// OTHER SOLUTION : With PIVOT
//============================
let
    Source = Excel.CurrentWorkbook(){[Name="tData059"]}[Content],
    FillUp_Seq = Table.FillUp(Table.ReplaceValue(Source,"#",null,Replacer.ReplaceValue,{"Seq"}),{"Seq"}),
    Group = Table.Group(FillUp_Seq, {"Seq"}, {{"All", each let
     A = Table.AddIndexColumn(_, "Option", 0),
     B = Table.FillDown( Table.AddColumn(A, "Q", each if ([Option]=0) then Text.From([Seq]) & "." & [Question] else null), {"Q"}),  
     C = Table.TransformColumns(Table.SelectRows(B, each [Option]<>0), { "Option", each "Option " & Text.From(_)})
     in C }})[[All]],
    Expand = Table.ExpandTableColumn(Group, "All", {"Q", "Option", "Question"}, {"Q", "Option", "Question"}),
    Pivot = Table.Pivot(Expand, List.Distinct(Expand[Q]), "Q", "Question"),
    NullRow = List.Repeat({null}, Table.ColumnCount(Pivot)),
    AnswerRow = {"Correct Answer"} & Table.SelectRows(Source, each [Correct] = "Y")[Question], 
    Result= Pivot & Table.FromRows( {NullRow} & {AnswerRow},Table.ColumnNames(Pivot))
in
    Result
*/