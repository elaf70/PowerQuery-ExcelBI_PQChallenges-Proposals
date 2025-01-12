/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7281519787770327040-dPLf

PQ proposal in 2 main steps
  (1) Transform rows with List.Accumulate
  (2) Back to Table

==================================*/

let
  Source = Excel.CurrentWorkbook(){[Name="tData250"]}[Content],
  TransformRows = List.Accumulate(Table.ToRows(Source), [ss=0, fs=0, r={}], 
    (s,c)=>[ss=c{2}??s[fs], fs=ss+c{3}-c{4}, r=s[r]&{{c{0},c{1},ss,c{3},c{4},fs}}]),
  Result = Table.FromRows(TransformRows[r], Table.ColumnNames(Source)& {"Finish Stock"})
in
  Result
