/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7144537173998854145-8LVg

PQ proposal in 4 main steps
==================================*/
let
    TS=#time(9,0,0), TE=#time(21,0,0), D=#duration(0,0,15,0), 
    Source = Excel.CurrentWorkbook(){[Name="tData142"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Start Time", type time}, {"End Time", type time}}),
    Shifts = List.Buffer(Table.ToRows(ChangeType[[Start Time], [End Time]])),
    Times = List.Generate(()=>TS, each _< TE, each _+D ),
    Transform = List.Transform(Times, each let 
        ts=_, te=ts+#duration(0,0,15,0), 
        nb=List.Count(List.Select(Shifts, each ((_{0}<ts and _{1}>ts) or (_{0}>=ts and _{0}<te))))
        in [Time=Time.ToText(ts)&"-"&Time.ToText(te), Count=nb] ),
    ToTable =  Table.FromRecords(Transform)
in
    ToTable


