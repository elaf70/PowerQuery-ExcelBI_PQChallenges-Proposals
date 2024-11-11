/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7027848082159484928-sT4l
PQ proposal in 4 main steps
1) Convert DateToNum 
2) GroupByStore 
        with x3 Sub-Steps as well A) Join with AllDates of Periods B) AddSeq ByCust C) Expand Cols 
3) Expand 
4) ChangeType [Date]
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData050"]}[Content],
    DateToNum = Table.TransformColumnTypes(Source,{{"Date", type number}}),
    Group = Table.Group(DateToNum, {"Store"}, {"Data", each let
        _AllDates = Table.FromColumns({{List.Min(_[Date])..List.Max(_[Date])}}, {"Date2"}),
        _Join = Table.Join(_AllDates, "Date2", _, "Date", JoinKind.LeftOuter),
        _FillDown = Table.FillDown(_Join, {"Date", "Customer"}),
        _AddSeqByGroup = Table.Group(_FillDown, {"Customer"}, {{"Data", each Table.AddIndexColumn(_, "Seq",1,1)}})[[Data]],
        _Expand = Table.ExpandTableColumn(_AddSeqByGroup, "Data", {"Date2", "Customer", "Seq"}, {"Date", "Customer", "Seq"})
        in _Expand}),
    Expand = Table.ExpandTableColumn(Group, "Data", {"Date", "Customer", "Seq"}),
    NumToDate = Table.TransformColumnTypes(Expand,{{"Date", type date}})
in NumToDate

