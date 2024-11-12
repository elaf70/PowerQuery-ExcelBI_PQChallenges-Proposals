/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7232959812286771200-vPok

PQ proposal in 4 main steps
    1) Split original tables in Group 
    2) Transform each into an inner-table adding column Group at 1st
    3) Sort resulting list with Comparer related to sum of [Income] of each
    4) Combine these sorted inner-tables 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData211"]}[Content],
    SplitCols = List.Split(Table.ToColumns(Source), 2),
    Transform = List.Transform(SplitCols, each let
      _Group = Text.End(_{0}{0},1),
      _T1 = Table.FromColumns({{null,"Group", _Group}} & _),
      _T2 = Table.PromoteHeaders(Table.Skip(_T1)),
      _T3 = Table.SelectRows(_T2, each [Name]<>null),
      _Fill = Table.FillDown(_T3, {"Group"}),
      _Sort = Table.Sort(_Fill, {"Income", Order.Descending}) 
      in _Sort ),
    Sort = List.Sort(Transform, (x,y)=>Value.Compare(List.Sum(y[Income]), List.Sum(x[Income]))),
    Combine = Table.Combine(Sort)
in
    Combine





