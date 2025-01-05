/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7281157410747039744-o5QC

PQ Proposal in 3 main steps
  (1) Group by [Class] with BestSudents calculation
  (2) Get ColumnNames from the inner table with the highest cols.count 
  (3) Expand inner grouped-Tables with it
==================================*/

let
  Source = Excel.CurrentWorkbook(){[Name="tData249"]}[Content],
  Group = Table.Group(Source, "Class", {"G", (t)=>let
    _CN = List.Transform({1..Table.RowCount(t)}, each "Student"&Text.From(_)) & {"Best Student"},
    _BS = Text.Combine(Table.SelectRows(t, each [Marks]=List.Sort(t[Marks], Order.Descending){0})[Student], ", ")
    in Table.FromRows({t[Student] & {_BS}}, _CN)  }),    
  CN = List.Max(List.Transform(Group[G], Table.ColumnNames),{},(x,y)=>Number.From(List.Count(x)>List.Count(y))),
  Expand = Table.ExpandTableColumn(Group, "G", CN)
in
  Expand
