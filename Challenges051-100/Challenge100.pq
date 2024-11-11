/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7091266157277515776-jpuE
PQ proposal in 2 main steps
    1) Group
    2) Expand
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData100"]}[Content],
    Group =Table.Group(Source , {"Name"}, {"Data", each let
        _R = List.Transform(Table.ToRecords(_), each let
            _Rate = _[Rate],
            _NbDays = Duration.Days(Date.From(_[To Date]??#date(2022,12,31))-Date.From(_[From Date])),
            _WD = List.Select(List.Dates(Date.From(_[From Date]), _NbDays+1, #duration(1,0,0,0)), 
                each Date.DayOfWeek(_, Day.Monday)<5)
            in List.Transform(_WD, each [Q="Q"&Text.From(Date.QuarterOfYear(_)), R=_Rate]) ),
        _T = Table.FromRecords(List.Combine(_R)),
        _Group = Table.Group(_T, "Q", {"V", each List.Sum(_[R])}) 
        in Table.Pivot(_Group, List.Distinct(_Group[Q]), "Q", "V", List.Sum)  }),
    Expand = Table.ExpandTableColumn(Group, "Data", {"Q1", "Q2", "Q3", "Q4"})
in
    Expand


