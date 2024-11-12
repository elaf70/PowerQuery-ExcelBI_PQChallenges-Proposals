/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7238395630518194176-u4rK

PQ proposal in 3 main steps
    1) ...
    2) ...
    3) ... 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData216"]}[Content],
    OriginalRows = List.Buffer(Table.ToRows(Source)),
    ColNames = List.Accumulate(OriginalRows, [i=0, r={}], (s,c)=>let
        _i = s[i]+List.Count(List.RemoveNulls(c)),
        _Name = "Items " & Number.ToText(s[i]+1) & " - " & Number.ToText(_i)
        in [i=_i, r=s[r] & {_Name}] )[r],
    Result = Table.FromColumns(OriginalRows, ColNames)
in
    Result


