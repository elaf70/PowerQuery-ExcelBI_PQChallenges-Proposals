/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7278983948016209920-j2Ze

PQ-proposal in 6 main steps 

NB: For FUN & to try not to use Unpivot+Pivot globally => do it instead with List functions mainly

==================================*/

let
  Source = Excel.CurrentWorkbook(){[Name="tData248"]}[Content],
  Cats = List.Distinct(Source[Category]),
  QT = Table.RemoveColumns(Source, {"Persons", "Category"}),
  NewC1 = List.Accumulate(List.RemoveNulls(Source[Persons]), {}, (s,c)=>s & {List.Last(Text.Split(c, ", "))}),
  NewQC = List.Transform( 
    List.Combine(List.Transform(Table.ToColumns(QT), each {List.Alternate(_,1,1,1), List.Alternate(_,1,1)})),
    each List.Accumulate(List.Skip(_), {_{0}}, (s,c)=>s &{c-List.Sum(s)}) ), 
  NewQCN = List.Accumulate(Table.ColumnNames(QT), {}, (s,c)=>s & List.Transform(Cats, each c & " " & _)),
  Result = Table.FromColumns({NewC1} & NewQC, {"Quarters"} & NewQCN)
in
  Result
