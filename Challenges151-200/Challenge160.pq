/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7167367606759362560-1Fzr

PQ proposal using List Accumulate twice to generate 1 CartesianProduct of data before distance calculation. 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData160"]}[Content],
    C = List.Buffer(Table.ToRows(Source)),
    Transform = List.Accumulate(C, {}, (s,c1)=> let
        _L = List.Accumulate(C, {c1{0}}, (s,c2)=>let
            _d = Number.Sqrt(Number.Power(c2{1}-c1{1},2)+Number.Power(c2{2}-c1{2},2))
            in s & {_d})
        in s & {_L} ),
    Result = Table.FromRows(Transform, {"Cities"} & Source[Cities])    
in
    Result


