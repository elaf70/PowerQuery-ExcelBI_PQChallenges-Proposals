/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7202521030186172416-lZi1

PQ proposal in 2 main steps  + few previous initializations values
    1) Group by [Continent] with addMissingContinent + Add Total Rows
    2) Combine previous Groupings and Add GrandTotal
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData187"]}[Content],
    CN = Table.ColumnNames(Source),
    CList = List.Sort(List.Distinct(Source[Continent])), 
    YList = List.Sort(List.Distinct(Source[Year])), 
    YRange = Text.From(List.First(YList)) & "-" & Text.From(List.Last(YList)),
    TotalSales = List.Sum(Source[Sales]),
    Group = Table.Group(Source, "Year", {"G", (t)=>let
        _Y = t[Year]{0},
        _MissingCRows = List.Transform(List.Difference(CList, t[Continent]), each {_, _Y, 0}),
        _AddMissingC = Table.FromRows( Table.ToRows(t) & _MissingCRows, CN),
        _Sort = Table.Sort(_AddMissingC, "Continent"),
        _Total = #table(CN, { {"TOTAL", _Y, List.Sum(t[Sales])}, {null,null,null}}) 
        in _Sort & _Total }),
    Add_GT = Table.Combine(Group[G]) & #table(CN, {{"GRAND TOTAL", YRange, TotalSales}}) 
in
    Add_GT


