/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7205414250305974273-Ynlk

PQ proposal  in 3 steps   =in fact 2 of previous definitions + 1 step only of transformation
    1) Define list of separators used for field values 
    2) Define a Splitter by CapitalLetter (not applying if no lower letter)
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData190"]}[Content],
    CN = {"Name:", "Org:", "City:", "FromDate:", "ToDate:"},
    fxSplitByCapitalSep = (x)=>if Text.PositionOfAny(x, {"a".."z"})=-1 then x 
       else Text.Combine(Splitter.SplitTextByCharacterTransition({"a".."z","A".."Z"},{"A".."Z"})(x), " "), 
    Split = Table.SplitColumn(Source, "Data", each 
      List.Transform(List.Skip(Splitter.SplitTextByAnyDelimiter(CN)(_)),
        each try Date.From(_) otherwise fxSplitByCapitalSep(_)),
      List.Transform(CN, each fxSplitByCapitalSep(Text.BeforeDelimiter(_,":"))) )
in
    Split



