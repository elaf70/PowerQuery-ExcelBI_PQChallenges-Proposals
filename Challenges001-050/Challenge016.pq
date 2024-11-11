/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6984725315071676416-haHm
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/

//SOLUTION 1 : Only with List functions - Without pivot/...
//---------------------------------------------------------
let
    Source = Excel.CurrentWorkbook(){[Name="tData016"]}[Content],
    PromoteHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    GrpYearColumns = List.Accumulate(
        List.Zip({  List.RemoveFirstN(Table.ColumnNames(PromoteHeaders),1) ,
                          List.RemoveFirstN(Record.ToList(Source{0})) }),
        {},
        (s,c) => if (c{0}=Text.From(c{1})) then s & {{c{0}}} 
                      else (List.RemoveLastN(s,1) & {List.Last(s) & {c{0}}})
    ),
    Result = List.Accumulate( GrpYearColumns, PromoteHeaders,
        (s,c)=>Table.CombineColumns(s, c, List.Sum, List.First(c))
    )
in
    Result

/*
//SOLUTION 2 : With Unpivot+pivot
//---------------------------------------------------------
let
    Source = Excel.CurrentWorkbook(){[Name="tData016"]}[Content],
    PromoteHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Unpivot = Table.UnpivotOtherColumns(PromoteHeaders, {"Sales Man"}, "Attribut", "Valeur"),
    ResetNull = Table.TransformColumns(Unpivot, {{"Attribut", each if Text.StartsWith(_, "Column") then null else _}}),
    FillDown = Table.FillDown(ResetNull,{"Attribut"}),
    Pivot = Table.Pivot(FillDown, List.Distinct(FillDown[Attribut]), "Attribut", "Valeur", List.Sum)
in
    Pivot
*/
