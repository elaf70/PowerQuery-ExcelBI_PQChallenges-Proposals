/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7222814447357087744-xIb5

PQ proposal in 3 main steps
    1) Create list of all subsets between 2 blank Amount1
    2) Use List.Accumulate to create for each with AmountValues 1 new {AvgAmount, GroupName} + Accumulate Remainings
    3) Create final result table with Table.FromRows 
==================================*/
let
    fxListAvgR = (l as list)=>Number.Round(List.Average(l)),
    Source = Excel.CurrentWorkbook(){[Name="tData203"]}[Content],
    Group = Table.Group(Source, "Amount1", {{"All", each _}}, GroupKind.Local, (x,y)=>Number.From(y=null)),
    Accumulate = List.Accumulate(Group[All],[g={}, r={}], (s,c)=>let
        _R = Table.ToRows(c),
        _Rem = List.Select(_R{0}, each _ is number),
        _Grp = List.Select(List.Combine(List.Skip(_R)), each _ is number)
        in [g=s[g] & List.RemoveNulls({fxListAvgR(_Grp)}), r=s[r] & _Rem]),
    GrpRows = List.Accumulate(Accumulate[g], {}, (s,c)=>s & {{"Group"&Text.From(List.Count(s)+1),c}}),     
    Result = Table.FromRows(GrpRows & {{"Remaining", fxListAvgR(Accumulate[r])}}, {"Group", "Avg Amount"})
in
    Result

