/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7182587900008488960-NtYQ

PQ proposal in 3 main steps
    1) Transform Lines in splitted RecordsList {[CommissionAmount, Agent]} with List.Accumulate
    2) Group by Agent to calculate each result
    3) Add GrandTotal
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData172"]}[Content],
    Transform = List.Transform(Table.ToRecords(Source), each let
        _C = [Amount]*[#"Commission %"]/100, 
        _AS = List.Zip({List.RemoveNulls({[Agent1],[Agent2]}), Text.Split([#"Share %"]??"100", ", ")})
        in List.Accumulate(_AS, {}, (s,c)=>s & {[A=c{0}, C=_C*Number.From(c{1})/100]}) ),
    Group = Table.Group(Table.FromRecords(List.Combine(Transform)), {"A"}, 
        {"C", each Number.Round(List.Sum([C]))}),
    Add_GT = Table.FromRows(Table.ToRows(Group) & {{"Total", List.Sum(Group[C])}}, {"Agent", "Commisssion"})
in 
    Add_GT


