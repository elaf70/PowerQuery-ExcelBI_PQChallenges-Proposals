/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7121706738004733952-Zpl9

PQ proposal in 1 main steps

==================================*/
let
    fxDateFrom = each Date.FromText(Text.From(_), [Format="yyyyMM"]),

    Source = Excel.CurrentWorkbook(){[Name="tData124"]}[Content],
    AddCol = Table.AddColumn(Source, "NetworkDayCount", each let
        _From = fxDateFrom([FromYearMonth]),
        _To = Date.EndOfMonth(fxDateFrom([ToYearMonth])),
        _Days = List.Select(
            List.Dates(_From, Duration.Days(_To - _From)+1, #duration(1,0,0,0)),
            each let 
                _DW = Date.DayOfWeek(_, Day.Monday),
                _WM = Date.WeekOfMonth(_, Day.Monday)
                in not(_DW=6 or (_DW=5 and List.Contains({1,3,5}, _WM))) )
        in List.Count(_Days) )
in
    AddCol


