/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7128954914591956992-X7Og

PQ proposal in 2 main steps
    1) ...
    2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData129"]}[Content],
    TData = Table.Buffer(Table.TransformColumnTypes(Source,{{"Date", type date}, {"Temp", Int64.Type}})),
    DRef = List.Max(TData[Date]),

    Transform = List.Transform( {7,30,365, "Month","Quarter","Year"}, (p)=>let
        DateFunctions = {Date.StartOfMonth, Date.StartOfQuarter,Date.StartOfYear}, 
        fxStartOf = (s)=>DateFunctions{List.PositionOf({"Month","Quarter","Year"},s)},
        _P = if (p is number) then "Last "&Text.From(p)&" Days" else p&" to Date", 
        _D = if (p is number) then Date.AddDays(DRef, -Number.From(p)+1) else fxStartOf(p)(DRef),
        _S = Table.SelectRows(TData, each [Date]>=_D),
        _ST=_S[Temp], _SD=_S[Date],
        _TMax = List.Max(_ST), _TMin = List.Min(_ST)
        in [Period=_P, 
            #"Max Temp"=_TMax, #"Max Temp Date"=_SD{List.PositionOf(_ST,_TMax)}, 
            #"Min Temp"=_TMin, #"Min Temp Date"=_SD{List.PositionOf(_ST,_TMin)}, 
            #"Avg Temp"=Number.Round(List.Average(_ST),2) ]),
    Result = Table.FromRecords(Transform)
in
    Result


