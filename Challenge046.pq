/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7022774849961869312-adwv
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData046"]}[Content],
    RI2M = [Month=1, Quarter=3, #"Half Yearly"=6, Year=12],   // Record To convert Interval=>Number of Month
    DFields = List.Transform({1..List.Max(Source[Period])}, (n)=>"Date"& Text.From(n)),
    Add_Dates = Table.AddColumn(Source, "Dates", (r)=> let 
        _Start = Date.From(r[Start Date]),
        _EOM = (_Start = Date.EndOfMonth(_Start)), 
        _Dates =  List.Transform({1..r[Period]}, (n)=>let
           _d = Date.AddMonths(_Start, n * Record.Field(RI2M, r[Interval]))
           in Date.ToText(if _EOM then Date.EndOfMonth(_d) else _d, "yyyy-MM-dd") )
        in Record.FromList(_Dates, List.FirstN(DFields, r[Period])) ),
    Expand = Table.ExpandRecordColumn(Add_Dates, "Dates", DFields)
in
    Expand
