/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7197445811071291393-pVwj

PQ proposal in 2 main steps
    1) Transform each row to a List of details rental values per Quarter
        using also List.Transform from related TotalPeriods index
    2) Combine To Table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData183"]}[Content],
    TransformRows = List.Transform(Table.ToRecords(Source), (r)=> let
        _DStart = Date.StartOfQuarter(#date(r[Year],3*Number.From(Text.End(r[Quarter],1)),1)),
        _IncY = 1 + (r[#"% Hike Yearly"] / 100)
        in  List.Transform({0..r[Total Periods]-1}, each let
            _Date = Date.AddQuarters(_DStart, _),
            _R = Number.Round(r[Rental] * Number.Power(_IncY, Number.IntegerDivide(_,4)))
            in {r[Vendor], Date.Year(_Date), "Q" & Text.From(Date.QuarterOfYear(_Date)), _R} ) ),
    Result = Table.FromRows(List.Combine(TransformRows), {"Vendor", "Year", "Quarter", "Rental"}) 
in
    Result


