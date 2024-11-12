/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7167005218356903936-5njT

PQ proposal in 2 main steps
    1) Group + List functions to build 1 inner table for all monthes 
    2) Expand
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData159"]}[Content],
    Group = Table.Group(Source, {"Name", "Year"}, 
        {"All", each let
            _r =  List.Zip({[Month],[Sales]}) & List.Transform(List.Difference({1..12},_[Month]), each {_,100})
            in Table.Sort(Table.FromRows(_r,  {"Month","Sales"}), "Month") }),
    Expand = Table.ExpandTableColumn(Group, "All", {"Month", "Sales"})
in
    Expand


