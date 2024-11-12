/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7159757463733387264-Ohks

PQ proposal in 2 main steps
    1) Group by [Pilot] and add EndOfMonth fictive flight stops 
        + calculate Flight an Rest times for each
    2) Group by Pilot/Year/Month to Sum 
==================================*/
let 
    fxSplitFlightByEOM = (fs as datetime, fe as datetime) as list=>let
        fxGen = (d as datetime)=>{d, List.Min({fe,Date.StartOfMonth(Date.AddMonths(d,1))})} 
        in List.Generate( ()=>fxGen(fs), each _{0}<fe, each fxGen(_{1}) ),
    Source = Excel.CurrentWorkbook(){[Name="tData154"]}[Content],
    GroupP = Table.Group(Source, "Pilot", {"All", each let
        SplitEOM = List.Transform(Table.ToRows(_), each fxSplitFlightByEOM (_{1},_{2})),
        LC = List.Combine(SplitEOM),
        Transform = let 
            fxGen = (i as number)=>
                [F=LC{i}, FS=F{0}, FE=F{1}, PFE=try LC{i-1}{1} otherwise FS, NFS=try LC{i+1}{0} otherwise FS, 
                EOM_FE=Date.EndOfMonth(FE), SOM_FS=Date.StartOfMonth(FS),
                Y = Date.Year(FS), M = Date.Month(FS), 
                FT = Duration.TotalHours(FE-FS),
              RT = Duration.TotalHours(if (NFS > EOM_FE) then (EOM_FE-FE) else (FS-List.Max({SOM_FS,PFE})))
//  To get same result as in challenge TOBE : but NOK -  incomplete calculations */                RT = Duration.TotalHours(if (NFS > EOM_FE) then (EOM_FE-FE) else (FS-PFE))
                ]
            in List.Generate(()=>[i=0] & fxGen(0), each [i]<List.Count(LC), each [i=[i]+1] & fxGen([i]+1)),
        TFR = Table.FromRecords(Transform, {"Y","M","FT","RT"}),
        GroupYM = Table.Group(TFR, {"Y","M"}, {{"FT", each List.Sum([FT])}, {"RT", each List.Sum([RT])}})   
        in GroupYM
    }),
    Expand = Table.ExpandTableColumn(GroupP, "All", {"Y","M","FT","RT"}, {"Year", "Month", "Fly Time", "Rest Time"})
in
    Expand

