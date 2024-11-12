/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7243469077933965313-hFh4

PQ proposal in 2 main steps + 1 first function definition
    0) 1 function to generale list of Text ShortMonth from 2 dates 
    1) Group by "Project" to generate 1 table by Month-columns for each 
    2) Expand all these inner-tables by project 
        with 1 global MonthList : created from Min(Start) and Max(Finsih) 
 ==================================*/
let
    fxListMonthName = (start as datetime,finish as datetime)=>let
        _LD = List.Generate(()=>Date.StartOfMonth(start), each _ <= finish, each Date.AddMonths(_,1) )
        in List.Transform(_LD, each DateTime.ToText(_, "MMM-yy")),

    Source = Excel.CurrentWorkbook(){[Name="tData220"]}[Content],
    Group = Table.Group(Source, "Project", {"G", (t)=> let
        _T1 = Table.TransformRows(t, each [[Activities]] &[Months=fxListMonthName([Start],[Finish])]),
        _T2 = Table.ExpandListColumn(Table.FromRecords(_T1), "Months"),
        _T3 = Table.Group(_T2, {"Months"}, {"G", each [Activities]}),
        _T4 = Table.FromColumns(_T3[G], _T3[Months])
        in _T4 } ),
    Expand = Table.ExpandTableColumn(Group, "G", 
        fxListMonthName(List.Min(Source[Start]), List.Max(Source[Finish])))
in
    Expand


/*
// SOLUTION 2 : List.TransformMAny
let
    Source = Excel.CurrentWorkbook(){[Name="tData220"]}[Content],
    Transform = Table.TransformColumns(Source, {
        {"Project", each _}, 
        {"Activities", each _}}, Date.StartOfMonth),
    I = List.Generate(() => List.Min(Transform[Start]), each _ <= List.Max(Transform[Finish]), each Date.AddMonths(_, 1)),
    T = List.TransformMany(List.Distinct(Transform[Project]), 
            each List.Zip(List.Transform(I, 
                    (d) => Table.SelectRows(Transform, (r) => r[Project] = _ and r[Start] <= d and d <= r[Finish])[Activities])), 
                (i, _) => {i} & _), 
    S = Table.FromRows(T, {"Project"} & List.Transform(I, each DateTime.ToText(_, "MMM-yy", "en-US")))
in
    S
*/

