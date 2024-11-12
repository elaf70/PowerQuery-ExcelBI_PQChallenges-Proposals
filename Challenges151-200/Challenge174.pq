/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7185124621128482816-NrF7

PQ proposal in 3 main steps
    1) Change Date fields type 
    2) Transform Rows with List.Generate to get for each,  List of expected NewRow-Values 
    3) Combine SubLists + Transform it into a Table 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData174"]}[Content],
    ChgType = Table.TransformColumnTypes(Source,{{"From Date", type date}, {"To Date", type date}}),
    Transform = Table.TransformRows(ChgType, each let
        E=[Emp], DStart= [From Date], DEnd = [To Date],         SalesPerDay = [Sales]/(Duration.Days(DEnd-DStart)+1),
        SDates = List.Generate(()=>DStart, each _<=DEnd, each Date.AddDays(Date.EndOfMonth(_),1)),
        NewRowsValues = List.Accumulate(SDates, [v={}, rt=0], (s,c)=>let
            _ds = c, _de = List.Min({Date.EndOfMonth(c), DEnd}),
            _MSales = (Duration.Days(_de - _ds)+1) * SalesPerDay, 
            _RT = s[rt] + _MSales
            in [v=s[v] & {{E, _ds, _de, _MSales, _RT}}, rt=_RT*Number.From(Date.Month(_ds)<>12)])
        in NewRowsValues[v]),    
    Result = Table.FromRows(List.Combine(Transform),             {"Emp","From Date", "To Date", "Monthly Sales", "Running Total"})
in
    Result


