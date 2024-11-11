/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6982186457884954624-iJT6
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/

//SOLUTION 1 : using LIST.GENERATE 
//--------------------------------
let
    Source = Excel.CurrentWorkbook(){[Name="tData014"]}[Content],
    _Steps=List.Buffer(Source[Step]),
    _NbItems=List.Count(_Steps),
    Transform = List.Generate(
        ()=>[i=0, iSerie=1, NewStep=_Steps{0}],
        each [i] < _NbItems,
        each [i = [i]+1, 
            iSerie = [iSerie] +  (if ([i]=0  or _Steps{i}>=_Steps{i-1}) then 0 else 1),
            NewStep = _Steps{i} & (if (iSerie<2) then Number.ToText(iSerie, " (0)")) ],
        each [NewStep] 
    ),
    Result = Table.FromColumns({Transform}, {"Step"})
in  
    Result

/*
//SOLUTION 2 : using LIST.ACCUMULATE
//----------------------------------
let
    Source = Excel.CurrentWorkbook(){[Name="tData014"]}[Content],
    // Use List.Accumulate to Transform 1 List in a NewList
    NewSteps = List.Accumulate(Source[Step],{}, 
            (s,c)=> let 
                // Get previous item of NewValueList => Split to get related {Value, SerieN°Suffix} 
                _PrevData = try Text.Split(List.Last(s)," ") otherwise {"Z","(0)"},
                _NewSuffix = if c >= _PrevData{0} then  _PrevData{1} 
                            else Number.ToText(Number.From(Text.Select(_PrevData{1}, {"0".."9"}))+1, "(0)")
            in s & {c & " " & _NewSuffix} 
    ),
    Result = Table.FromList(
        List.Transform(NewSteps, each Text.Replace(_," (1)","")),
        null, {"Step"})
in
    Result
*/
