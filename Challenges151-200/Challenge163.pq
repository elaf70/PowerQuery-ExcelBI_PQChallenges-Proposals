/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7172078655110688768-bMX0

PQ proposal in 2 main steps
    1) Buffer the list of all VehiculeCodes
    2) Add new col. using List.Accumulate to generate VehiculeNumbers 
==================================*/
let
    Source = each Excel.CurrentWorkbook(){[Name = _]}[Content],
    LVC = List.Buffer(Source("tData163a")[Vehicle code]),
    Add_V = Table.AddColumn(Source("tData163b"), "Vehicule Numbers", each let
        S = Text.Replace(Text.Upper([Data]), " ", ""), 
        LVN = List.Accumulate({0..Text.Length(S)-10-1}, {}, (s,c)=>let
            VN = Text.Range(S, c, 10),
            Parts = Splitter.SplitTextByLengths({2,2,2,4})(VN),
            VN_OK = List.Contains(LVC, Parts{0}) and Parts{1}<>"00"                              and (Text.Remove(VN, {"0".."9"}) = Parts{0} & Parts{2})
            in if (VN_OK) then s & {VN} else s )    
        in Text.Combine(LVN, ", ") )
in
    Add_V



