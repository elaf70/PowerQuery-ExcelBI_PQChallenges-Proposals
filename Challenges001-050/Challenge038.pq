/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7012612536982237184-5_49
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData38"]}[Content],
    Transform =List.Accumulate(List.Skip(Source[Data]), {Source[Data]{0}}, 
           (s,c)=> s & {c &List.Last(s) & c}), 
    Result = Table.FromColumns({Transform}, {"Triangle"}) 
in
    Result

