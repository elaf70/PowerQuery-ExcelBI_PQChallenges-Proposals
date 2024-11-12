/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7245643396243963904-LSsf

PQ proposal in 1 main steps as a call to a recursive function to AddIndex

Of course the code is a little less "intuitive" at first sight this way 
But was really FUN searching for it ... 😃 (=1st time I ever try using Recurcivity in PQ) 
==================================*/

// SOLUTION 2 : Recursif
//=======================
let
    fxAddIndex = (t as table, cNames as list)=>let
        cn = cNames{0},
        lastIdx = List.Last(List.Select(Table.ColumnNames(t), each Text.EndsWith(_,"_Index"))),
        prefix = if (lastIdx=null) then "" else Table.Column(t,lastIdx){0}&".",
        values = List.Distinct(Table.Column(t,cn)),
        Add_Idx = Table.AddColumn(t, cn & "_Index", each prefix & Text.From(List.PositionOf(values, Record.Field(_,cn))+1))
        in  if (List.Count(cNames)>1) 
            then Table.Combine(Table.Group(Add_Idx, cn, {"G", each @fxAddIndex(_, List.Skip(cNames))})[G])
            else Add_Idx,

    Source = Excel.CurrentWorkbook(){[Name="tData221"]}[Content],
    Result =  fxAddIndex(Source, Table.ColumnNames(Source))
in 
    Result


/*
// SOLUTION 1 : Itératif
//=======================
let
    fxAddIndex = (t as table, cn as text, pIdx as text)=>let
        _Values = List.Distinct(Table.Column(t,cn))
        in Table.AddColumn(t, cn & "_Index", each pIdx & Text.From(List.PositionOf(_Values, Record.Field(_,cn))+1)),
    Source = Excel.CurrentWorkbook(){[Name="tData221"]}[Content],
    Add_PIdx = fxAddIndex(Source,"Project", ""), 
    GroupP = Table.Group(Add_PIdx, "Project", {"G", each let
        Add_TIdx= fxAddIndex(_, "Task", [Project_Index]{0}&"."),
        GroupT = Table.Group(Add_TIdx, "Task", {"G", each fxAddIndex(_, "Activity",[Task_Index]{0}&".")})
        in Table.Combine(GroupT[G]) }),
    Result = Table.Combine(GroupP[G])
in
    Result
*/