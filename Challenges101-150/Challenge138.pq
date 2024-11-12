/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7139463742718746624-9vTx

PQ proposal in 4 main steps with  GroupSize that can be set as parameter
    1) ...
    2) ...
    3) ... 
    4) ...
==================================*/
let
    NGrp = 2,
    CN = List.Combine(List.Accumulate({1..NGrp}, {{},{}}, 
        (s,c)=> {s{0} & {"Group"&Text.From(c)}, s{1} & {"Value"&Text.From(c)}})),    

    Source = Excel.CurrentWorkbook(){[Name="tData138"]}[Content],
    ToRows = List.Transform(Table.ToRows(Source), List.RemoveNulls),
    All = List.Zip({List.Combine(List.Alternate(ToRows,1,1,1)),List.Combine(List.Alternate(ToRows,1,1,0))}),
    Transform = List.Transform(List.Split(All,NGrp), 
        (_GVs as list)=>List.Combine(List.Accumulate({0..NGrp-1}, {{},{}}, 
            (s,c)=>{s{0}&{try _GVs{c}{0} otherwise null},s{1}&{try _GVs{c}{1} otherwise null}}) )),
    ToTable = Table.FromRows(Transform, CN)
in
    ToTable


