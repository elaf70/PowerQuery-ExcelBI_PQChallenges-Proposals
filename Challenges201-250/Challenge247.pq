/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7278620688741990400-ENh6

PQ-proposal with following 5 main steps
  (1) Join 2 source tables (2) Check Questions result 
  (3+4) Rename Question column and Pivot (5) Add row scores

==================================*/

let
  fxSource = (n)=>Excel.CurrentWorkbook(){[Name=n]}[Content],
  Join = Table.Join(fxSource("tData247_1"), "Question", fxSource("tData247_2"), "Question", JoinKind.Inner),
  CR = Table.CombineColumns(Join,{"Option Chosen", "Correct Option"},(x)=>if (x{0}=x{1}) then "Y" else "N","R"), 
  TQ= Table.TransformColumns(CR, {"Question", each "Q"&Text.From(_)}),
  Pivot = Table.Pivot(TQ, List.Distinct(TQ[Question]), "Question", "R", each _{0}),
  AddScore = Table.AddColumn(Pivot, "Score", each List.Count(List.Select(Record.FieldValues(_), each _="Y")))
in
  AddScore
