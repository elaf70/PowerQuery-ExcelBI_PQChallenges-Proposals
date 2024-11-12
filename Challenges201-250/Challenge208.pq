/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7228249881482600448-yXpt

PQ proposal in 2 main steps
    1) Group
    2) Combine
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData208"]}[Content],
    Group = Table.Group(Source, "Month", {"G", each let
        _mv = List.RemoveNulls([3 Year MV]),
        _Result = List.Accumulate(_mv, {_mv{0},_mv{0}}, (s,c)=>s & {c*3 - List.Sum(List.LastN(s,2))}) & {0}
        in Table.FromColumns({[Month], [Year], _Result}, {"Month", "Year", "Defect"})
    }),
    Combine = Table.Combine(Group[G]) 
in
    Combine



