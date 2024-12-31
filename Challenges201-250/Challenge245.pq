/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7276083980175577088-0Lfx

PQ proposal in 6 main steps
  Didn't succeed to find anything shorter today = 🤔   
   => Eager to look for other proposals

==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData245"]}[Content],
  Transform = Table.FromRows(
    List.Accumulate(Table.ToRows(Source), {}, (s,c)=>s & List.Transform(Text.Split(c{1},", "), each {c{0}, _})), {"Name", "Subject"}),
  Group = Table.Group(Transform, "Subject", {"G", each List.Repeat(List.Sort(_[Name]), Number.RoundUp(5 / List.Count(_[Name])))}),
  Sort = Table.Sort(Group, "Subject"),
  NbBackup = List.Max(List.Transform(Sort[G], List.Count))-5,
  Backups = List.Transform({1..NbBackup}, each "Backup" & Text.From(_)),
  ToTable = Table.FromColumns( {{"Mond","Tue","Wed","Thu","Fri"} & Backups} & Sort[G], {"Days"} & Sort[Subject])
in
  ToTable
