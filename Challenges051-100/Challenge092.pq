/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7081119296285794304-AO13
PQ proposal in 3 main steps
    1) ...
    2) ...
    3) ... 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData092"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Add_Month = Table.AddColumn(ChangeType, "Month", each let
        _EOM = Date.EndOfMonth([Date]),
        _LastFriday = Date.AddDays(_EOM, -Date.DayOfWeek(_EOM,Day.Friday)*Number.From(Date.Month(_EOM)<>12)),
        _Month = #date(1900,Date.Month([Date]) + Number.From([Date]>_LastFriday),1)
        in Text.Start(Date.MonthName(_Month,"en-US"),3) ),
    Group = Table.Group(Add_Month, "Month", {"Total Sales", each List.Sum([Sales])})
in
    Group


