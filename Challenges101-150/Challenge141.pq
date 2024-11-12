/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7144176565151600640-CoaZ

PQ proposal in 2 main steps
    (1) Group per month to add 2 colums of Moving Average 
        each calculated with 1 function template from List.Accumulate 
    (2) Combine previous Grouping result 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData141"]}[Content],
    CN = {"Year", "3 Year MV", "5 Year MV"}, 
    Group = Table.Group(Source, {"Month"}, {"All", each let
        fxListTransformMV = (l, p)=> List.Accumulate({1.. List.Count(l)}, {}, 
            (s,c)=> s & {if (c<=p) then null else Number.Round(List.Average(List.Range(l, c-p-1, p)))} ),
        _Y = [Year],
        _D = List.Buffer([Defects]),
        _3YMV = fxListTransformMV(_D, 3),
        _5YMV = fxListTransformMV(_D, 5)
        in Table.FromColumns({_Y, _3YMV, _5YMV}, CN) }),
    Expand  = Table.ExpandTableColumn(Group, "All", CN)
in
    Expand











