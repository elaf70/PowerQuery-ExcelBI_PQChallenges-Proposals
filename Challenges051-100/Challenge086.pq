/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7073509152483946496-R7BB
PQ proposal 
    with LIST manipulations only 
    + try to reduce quantity of calculations needed as well
    For FUN & also in search of a less "classical" approach 
==================================*/
let
    T1 = Excel.CurrentWorkbook(){[Name="tData086_1"]}[Content],
    T2 = Excel.CurrentWorkbook(){[Name="tData086_2"]}[Content],
    // Preparation
    Names = T1[Name], 
    TRanges = Table.SelectRows(T2, each List.Contains(Names, [Name]) ),
    DMin = List.Min(TRanges[From]), DMax = List.Max(TRanges[To]),
    // Calculate Columns Date
    CDates = {Number.From(DMin)..Number.From(DMax)},
    // Calculate each Name-column with Repetitions of Y/N 
    // from NbDays Before/Between/After each periode 
    CData = List.Accumulate(Names, {}, (s,c)=>let 
        _D = Table.SelectRows(TRanges, each [Name]=c){0}?,
        _Col = if (_D=null) then List.Repeat({"N"}, Duration.Days(DMax-DMin)+1)
            else List.Repeat({"N"}, Duration.Days(_D[From]-DMin)*Number.From(_D[From]>DMin)) 
            & List.Repeat({"Y"}, Duration.Days(_D[To]-_D[From])+1)
            & List.Repeat({"N"}, Duration.Days(DMax-_D[To])*Number.From(DMax>_D[To]))
        in s &{_Col}),
    // Result = Table FromColumns +ChangeType for Date       
    Result = Table.FromColumns({CDates} & CData, {"Dates"} & Names), 
    ChangeType = Table.TransformColumnTypes(Result,{{"Dates", type date}})
in
    ChangeType


