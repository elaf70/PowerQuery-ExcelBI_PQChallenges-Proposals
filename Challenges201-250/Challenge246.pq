/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7276446357794570240-MeCX

PQ-proposal in 3 main steps
 1) Group by [Deal ID] + Transform 
 2) Combine previous inner tables 
 3) Pivot the [Designation] column as a separated global step => to get related columns-order as required  
    NB : Had it first, included in 1st Group step in a 1st release : but without this proper order        
==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData246"]}[Content],
  Group = Table.Group(Source, "Deal ID", {"G", (t)=>let
    _DMax = List.Max(t[Date])
    in Table.SelectRows(t, each ([Date] = _DMax))  }),
  Combine = Table.RemoveColumns(Table.Combine(Group[G]), {"Date"}),      
  Pivot = Table.Pivot(Combine, {"Mgr", "GM", "VP"}, "Designation", "Name", each Text.Combine(_, ", "))
in
  Pivot
