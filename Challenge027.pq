/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-6999587751394226176-SVyN
PQ proposal in 3 main steps
1) Get sorted List all of [Club] 
2) Generate all matches records with 2 list.Acccumulate
3) Convert it as final result Table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData027"]}[Content],
    Clubs = List.Sort(Source[Clubs]),
    Matchs = List.Accumulate(Clubs, [R={}, C=Clubs], (s,c)=> let 
        _OtherClubs = List.RemoveItems(s[C], {c}),
        _Matchs = List.Accumulate(_OtherClubs,{}, (s2,c2)=>s2&{[Team1=c, Team2=c2]})
        in [R=s[R] & _Matchs, C=_OtherClubs])[R],
    ToTable = Table.FromRecords(Matchs)
in
    ToTable

