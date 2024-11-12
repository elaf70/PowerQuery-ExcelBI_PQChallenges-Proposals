/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7149610605153337344-vlll

PQ proposal in 2 main steps
    1) GroupBy and select from each [Group] only matching rows by a Table.SelectRows call
    2) Combine inner-tables 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData146"]}[Content],
    Group = Table.Group(Source, "Group", {"All", each let
        T = _{0}[Threshold], V = _[Value],
        PTV = List.Last(List.Select(V, each _<T)),
        NTV = List.First(List.Select(V, each _>T))
        in Table.SelectRows(_, each List.Contains({PTV, NTV}, [Value])) }),
    Result = Table.Combine(Group[All])
in
    Result


