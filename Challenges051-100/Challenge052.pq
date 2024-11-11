/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7012265601683337216-vs0P
PQ proposal in 9 main steps

I was looking in a way to keep same algorithm for even/odd clubs. 
=Mainly based on ListAccumulate to create records for each match
       + get it finally (I hope) also for even nb of clubs with a simple 
       with a roundup of Nb to extract + a check of "duplicate" match(*)
(*) Not always needed but you can apply it same way on both case

==================================*/
 let
    Source = Excel.CurrentWorkbook(){[Name="tData052"]}[Content],
    Clubs = List.Sort(Source[Clubs]),
    Clubsx2 = List.Combine({Clubs,Clubs}),
    NC= List.Count(Clubs),
    NCDiv2 = Number.IntegerDivide(NC,2),
    NCMod2 = Number.Mod(NC,2),

    Start = Table.FromColumns({Clubs,{1..NC}},{"Team 1", "Team 2"}),
    Transform = Table.TransformColumns(Start, 
        {"Team 2", (i)=>List.Range(Clubsx2, i, NCDiv2-Number.From(i>NCDiv2 and NCMod2=0))} ),
    Expand = Table.ExpandListColumn(Transform, "Team 2"),
    Sort = Table.Sort(Expand, {{"Team 1", Order.Ascending}, {"Team 2", Order.Ascending}})
in 
    Sort

