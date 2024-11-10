/*===============================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6967339867827503104-GDH8
PQ proposal in 2 main-steps
1) Transform each row by splitting LeaveDates by 2 and or each find days then count the combination of all distinct ones 
2) Sort by [Employee]
=================================*/

SOLUTION 1 :   Get all digits (even if it appears several times in the string)
============
let
  Source = Excel.CurrentWorkbook(){[Name="tData001"]}[Content],
  Transform = Table.TransformRows(Source, each let
    _S = Text.From([String]),
    _LD = List.Select(Text.ToList(_S), each List.Contains({"0".."9"}, _)),
    _D = Text.Combine(_LD,""),
    _SD = List.Sum(List.Transform(_LD, each Number.FromText(_)))
    in {_S, _D, _SD} ),
  Result = Table.FromRows(Transform, {"String", "Digits", "Sum Of Digits"})
in 
  Result

SOLUTION 2 : With each digit only detected once
============
let
  Source = Excel.CurrentWorkbook(){[Name="tData001"]}[Content],
  Transform = Table.TransformRows(Source, each let
    _S = Text.From([String]),
    _LD = List.Intersect({Text.ToList(_S), {"0".."9"}) ),
    _D = Text.Combine(_LD,""),
    _SD = List.Sum(List.Transform(_LD, each Number.FromText(_)))
    in {_S, _D, _SD} ),
  Result = Table.FromRows(Transform, {"String", "Digits", "Sum Of Digits"})
in 
  Result
