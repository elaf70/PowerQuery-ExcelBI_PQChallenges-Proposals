/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7190226035232174080-HNW-

PQ proposal in 3 main steps : Transform + Select + Combine
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData178"]}[Content],
    Transform = List.Transform(Table.ToRows(Source), each let
        L = {{_{0}, "Role", _{1}, _{2}}, {_{0}, "Org", _{3}, _{4}}}
        in List.Select(L, each _{3}<>null)), 
    Result = Table.FromRows(List.Combine(Transform), {"Emp","Change","Old","New"})
in
    Result


