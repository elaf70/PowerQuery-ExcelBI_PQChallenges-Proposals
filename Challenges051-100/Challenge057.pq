/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7037632752900317184-Av_N
PQ proposal in 3 main steps
    1) GroupBy [Group] and determine MaxSeqNumber
    2) List.Accumulate with nested List.Generate for Seq. order to generate all records
    3) To build with it, the final result table
==================================*/

let
    Source = Excel.CurrentWorkbook(){[Name="tData57"]}[Content],
    Group = Table.Group(Source, {"Group"}, {{"All", each _[Group]}})[All],
    MaxSeqItems = List.Max(List.Transform(Group, List.Count)),
    Transform = List.Accumulate( Group, [Index=1,List={}], 
        (s,c)=>let
            _LC=List.Count(c),
            Records = List.Generate(
                ()=>[x=1, y=[Seq=Text.From(s[Index]), Group=c{0}]], 
                each [x] <= MaxSeqItems,
                each let
                    _NewSeq = if ([x]<_LC) 
                        then Text.From(s[Index]+[x]) 
                        else (Text.From(s[Index]+_LC-1) & "." & Text.From([x]-_LC+1))
                    in [ x=[x]+1, y=[Seq=_NewSeq, Group=c{0}] ],
                each [y])
            in [Index=s[Index]+_LC, List=List.Combine({s[List],Records})] 
        ),
    Result=Table.FromRecords(Transform[List])
in
    Result

/*
// SOLUTION 2 : try to make the seq geenration a little more simple (: not sure this really does it ?)
let
    Source = Excel.CurrentWorkbook(){[Name="tData057"]}[Content],
    Group = Table.Group(Source, {"Group"}, {{"SeqNb", each List.Count(_[Seq])}}),
    MaxSeqItems = List.Max(Group[SeqNb]),
    StartList = List.RemoveLastN(List.Accumulate(Group[SeqNb], {1}, (s,c)=>s&{List.Last(s)+c}),1),
    Add_Start = Table.FromColumns(Table.ToColumns(Group) & {StartList}, {"Group","SeqNb","Start"}),
    Add_Seq = Table.AddColumn(Add_Start, "Seq", each
        let 
            n1=[Start]+[SeqNb]-1,  n2=MaxSeqItems-[SeqNb],
            l1=List.Transform({[Start]..n1},Text.From),
            l2=List.Transform(List.Zip({List.Repeat({n1},n2), {1..n2}}), 
                each Text.From(_{0})&"."&Text.From(_{1}))
        in l1 & l2 ),
    Expand = Table.ExpandListColumn(Add_Seq, "Seq")[[Seq],[Group]] 
in
    Expand
*/

