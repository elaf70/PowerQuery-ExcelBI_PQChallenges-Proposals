/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7218101919368527872-2XEd

PQ proposal in 3 main steps
    1) Combine
    2) Unpivot
    3) Pivot
==================================*/
let
    Source = Table.SelectRows(Excel.CurrentWorkbook(), each Text.StartsWith([Name],"tData200"))[Content],
    Combine = Table.Combine(Source),
    Unpivot = Table.UnpivotOtherColumns(Combine, {"Student"}, "Attr", "Val"),
    Pivot = Table.Pivot(Unpivot, List.Sort(List.Distinct(Unpivot[Attr])), "Attr", "Val", List.Max)
in
    Pivot


