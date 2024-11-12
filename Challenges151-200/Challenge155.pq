/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7161931798501867520-G4EA

PQ proposal in 2 main steps
    1) List.Accumulate to select all substrings "like" "[h]h:m[m]"
    2) List.Select to keep only those as proper time value 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData155"]}[Content],
    Result = Table.AddColumn(Source, "Hours", each let 
        l = List.Generate(()=>[s=Text.Select([String], {"0".."9",":"})&" ", r={}], 
            each Text.Length([s])>0,
            each let
                h = Text.End(Text.BeforeDelimiter([s],":"),2),
                m = Text.Start(Text.AfterDelimiter([s],":"),2),
                hm = if (Text.Contains([s],":")) then h & ":" & m else [s] 
                in [s=Text.AfterDelimiter([s], hm), r=[r] & {hm}]
            ),            
        s = List.Select(List.Last(l)[r], each try Time.From(_) is time otherwise false) 
        in  Text.Combine(s, ", ")
        )
in
    Result


