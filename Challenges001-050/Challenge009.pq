/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6976750557134774272-6Hyy
PQ proposal in 3 main steps
1) ...
2) ...
3) ... 
==================================*/
let
    fxGetXLTable = (tName)=> Excel.CurrentWorkbook(){[Name=tName]}[Content],
    LFromTo = List.Buffer(Table.ToRows(fxGetXLTable("tData009b"))),
    Separators=" ,'." , 
    Source = fxGetXLTable("tData009"),
    NewText = List.Transform(Source[Original Text], each let
        _Words = List.ReplaceMatchingItems(Text.SplitAny(_, Separators), LFromTo),  
        _Sep = Text.ToList(Text.Select(_, Text.ToList(Separators))),
        _Text = List.Accumulate({1..List.Count(_Words)-1}, _Words{0}, (s,i)=>s & _Sep{i-1} & _Words{i})   
        in _Text),
    Result = Table.FromColumns({NewText},{"Final Text"})
in
    Result
