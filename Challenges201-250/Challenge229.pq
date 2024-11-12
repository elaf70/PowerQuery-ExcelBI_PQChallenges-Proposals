/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7255803394009792512-_ff1

PQ proposal in 2 main steps
    1) Add as new column for each row 1 inner table of all amounts per YearMonth of the period, splitting it with List.Generate 
    2) Expand it 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData229"]}[Content],
    NewColNames = {"YearMonth", "Amount"},
    Add_Split = Table.AddColumn(Source, "Split", each let
        _DStart = Date.From([From Date]), 
        _DEnd = Date.From([To Date]) ,
        _AmountPerDay = [Amount] / (Duration.Days(_DEnd-_DStart)+1),  
        _R = List.Generate(
            ()=>_DStart, each _<=Date.EndOfMonth(_DEnd), 
            each Date.StartOfMonth(Date.AddMonths(_,1)), 
            each {
                 Date.ToText(_,"MM-yy"), 
                 Number.Round(_AmountPerDay * (Duration.Days(List.Min({_DEnd, Date.EndOfMonth(_)}) - _)+1))
            } ) 
        in Table.FromRows(_R, NewColNames)),
    Expand = Table.ExpandTableColumn(Add_Split[[Transaction],[Split]], "Split", NewColNames)
in
    Expand


