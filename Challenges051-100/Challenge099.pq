/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7090903773715267584-Yn4V
PQ proposal in 3 main steps
    1) Change
    2) Add column [Month]
    3) Group by [Month] and calculate for each the 4 required columns 
==================================*/
let
    fxGetDates = (t, v)=>let 
          _LD = Table.SelectRows(t, each [Amount]=v) 
          in Text.Combine(List.Transform(_LD[Date], each Date.ToText(_,"yyyy-MM-dd")),", "),

    Source = Excel.CurrentWorkbook(){[Name="tData099"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source, {{"Date", type date}, {"Amount", Int64.Type}}),
    Add_Month = Table.AddColumn(ChangeType, "Month",          each Date.ToText([Date],"MMM"), type text),
    Group = Table.Group(Add_Month, {"Month"}, {
        {"Min", each List.Min(_[Amount])},
        {"Min Date", each fxGetDates(_, List.Min(_[Amount]))},
        {"Max", each List.Max(_[Amount])},
        {"Max Date", each fxGetDates(_, List.Max(_[Amount]))}
        })
in
    Group


