/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7124241956502667264-QIs4

PQ proposal in 2 main steps
    1) ChangeType of [Date]
    2) Add New column with splitting [Reference] to determine how many days to add to [Date] to get the expected result 
 ==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData126"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Add_RD = Table.AddColumn(ChangeType, "Reference Date", each let
     _s = Splitter.SplitTextByAnyDelimiter({"st ", "nd ", "rd ", "th "})([Reference]),
     _dw1 = Date.DayOfWeek([Date], Day.Monday),
     _dw2 = List.PositionOf({"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"}, _s{1}),
     _d = Date.AddDays([Date], _dw2 - _dw1 + 7*(Number.From(_s{0}) - Number.From(_dw2>_dw1)))
     in _d, type date)
in
    Add_RD


