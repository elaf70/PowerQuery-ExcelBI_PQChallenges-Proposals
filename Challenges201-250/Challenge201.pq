/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7220272337693122560-Mpkh

PQ proposal in 4 main steps
    1) Add column to [Buyer] table with days list of buying period 
    2) Add column to [Item] table with days list of stock period. Having Min/Max of all buyers Period if 1 stock date is missing
    3) For each buyer buid a record for all items with 
        NbOfDaysInCommun of 2 periods buying & stock 
        Calculated from List.Intersect of days List of the 2 periods
    4) Final Result = table from these calculated records
==================================*/
let
    Source = Table.SelectRows(Excel.CurrentWorkbook(), each Text.StartsWith([Name],"tData201"))[Content],
    fxDates = (s as date,e as date)=>List.Dates(s, if(e>=s) then Duration.Days(e-s)+1 else 0, #duration(1,0,0,0)),
    B = Table.AddColumn(
        Table.TransformColumnTypes(Source{0},{{"Buy Date From", type date}, {"Buy Date To", type date}}),
        "Period", each fxDates([Buy Date From], [Buy Date To]) ),
    DMin = List.Min(B[#"Buy Date From"]),
    DMax = List.Max(B[#"Buy Date To"]),  
    I = Table.AddColumn(
        Table.TransformColumnTypes(Source{1},{{"Stock Start Date", type date}, {"Stock Finish Date", type date}}),
        "Period", each fxDates([Stock Start Date]??DMin, [Stock Finish Date]??DMax) ),
    Transform = Table.TransformRows(B, each let 
        B = [Buyer], BPeriod=[Period],
        Add_In = Table.AddColumn(I, "DaysIn", each List.Count(List.Intersect({[Period], BPeriod}))),
        R = List.Accumulate(Table.ToRecords(Add_In), [#"Buyer / Items" =B], 
            (s,c)=>s & Record.FromList({if c[DaysIn]>0 then "x" else null}, {c[Items]}))
        in R),
    Result = Table.FromRecords(Transform)
in
    Result








