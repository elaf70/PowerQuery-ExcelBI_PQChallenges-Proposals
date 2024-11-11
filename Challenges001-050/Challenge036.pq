/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7010091274472972288-nDxv
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/
let
    T1 = Excel.CurrentWorkbook(){[Name="tData036_1"]}[Content],
    Source = Excel.CurrentWorkbook(){[Name="tData036_2"]}[Content],
    Result = Table.AddColumn(Source, "Emp Name", each let 
       _P = Text.Split([Emp ID], "*"),
       _N = Text.Remove([Emp ID], "*"),
       _Fx = if (List.Count(_P)<>2) then Text.Contains else (if (_P{0}<>"") then Text.StartsWith else Text.EndsWith),
       _S = Table.SelectRows(T1, each _Fx(Text.From([Emp ID]), _N) )
       in  Text.Combine(_S[Emp Name], ", ") )
in
    Result

