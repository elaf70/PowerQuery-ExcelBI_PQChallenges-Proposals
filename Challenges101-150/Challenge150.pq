/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7154684033640103936-yJ-n

PQ proposal in 3 main steps
    1) Store ListOf NoNull TimeOut
    2) Transform Table Records with List.Accumulate
    3) Back to Table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData150"]}[Content],
    LTO = List.Buffer(List.RemoveNulls(Source[Time Out])),
    Transform = List.Accumulate(Table.ToRecords(Source), [r={}, i=0], (s,c)=>let  
        n = if ((c[Time In]<>null)) then c  
            else  Record.TransformFields(c, {{"Time In", each LTO{s[i]}}, {"Time Out", each LTO{s[i]}+c[Duration]/24}}) 
        in [r=s[r]&{n}, i=s[i]+Number.From(c<>n)]),
    ToTable = Table.FromRecords(Transform[r])
in
    ToTable


