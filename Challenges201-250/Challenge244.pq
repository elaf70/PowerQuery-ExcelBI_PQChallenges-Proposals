/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7273909647374196736-LjLw

"Late" PQ-Proposal for this challenge missed at first publication in 2 main steps
      (1) Combine 2 original tables
      (2) Group by [Student] with  required-sorted colums + group columns functions-calls calclulated dymanically  
          NB : rows sort not needed with this challenge-data 

For FUN : try first without any globalcalls of Unpivot/PIvot that first comes in mind. 
NB: Add it  later as 2nd proposal, with same logic/process   

==================================*/
let
  fxSource = (n)=>Excel.CurrentWorkbook(){[Name=n]}[Content],
  Combine = fxSource("tData244_1") & fxSource("tData244_2"),  
  Group = Table.Group(Combine, {"Student"}, List.Transform(List.Sort(List.Skip(Table.ColumnNames(Combine))), (c)=>{c, each List.Sum(Table.Column(_,c))}))
in
  Group


/*
//==============================================================
//  2nd proposal : with Unpivot & pivot + same logic/process for other requirements
//==============================================================
let
  fxSource = (n)=>Excel.CurrentWorkbook(){[Name=n]}[Content],
  Combine = fxSource("tData244_1") & fxSource("tData244_2"),  
  Unpivot = Table.UnpivotOtherColumns(Combine, {"Student"}, "A", "V"),
  Pivot = Table.Pivot(Unpivot, List.Sort(List.Distinct(Unpivot[A])), "A", "V", List.Sum)
in
  Pivot
