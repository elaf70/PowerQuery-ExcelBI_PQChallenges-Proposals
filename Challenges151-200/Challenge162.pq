/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7169904320287555584-IBcD

PQ proposal in 1 main steps

NB : Would like to have it more simple with RegularExp pattern. But don't think it's possible for now in M =>❓ 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData162"]}[Content],
    Result = Table.AddColumn(Source, "Result", each let
        s = [String],
        l = List.Select(Text.SplitAny(s, Text.Remove(s, {"0".."9"})), each _<>""),
        r = List.Accumulate({0..List.Count(l)-1}, [str=s, r={}], (s,c)=>let
            sNum = l{c}, sTxt = Text.BeforeDelimiter(s[str], sNum),
            _r = if (Text.Length(sTxt)<2 or Text.Length(sNum)<2) then {} 
                else let
                    sC1 = Text.End(sTxt,1),
                    sC2 = Text.Range(sTxt, Text.Length(sTxt)-2,1)
                    in  if (not(List.Contains({"A".."Z"},sC1)) and List.Contains({"A".."Z"},sC2)) 
                        then {sC2 & Text.Start(sNum,2)} else {}
            in [str=Text.AfterDelimiter(s[str], sNum), r=s[r] & _r])
        in Text.Combine(r[r],", ") )
in
    Result



