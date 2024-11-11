/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6972039705718632448-5jhr
PQ proposal in 4 main steps
     (1)CleanDistinctSubjects, (2)Expand, (3)Group-Count, (4)Sort  
==================================*/

let
    _NormalChar = {"A".."Z", "0".."9"},
    Source = Excel.CurrentWorkbook(){[Name="tData006"]}[Content],
    CleanSubjectsList = Table.TransformColumns(Source, {
      {"Subjects", each let 
          _CharList = Text.ToList(Text.Upper(_)),
          _ReplaceSep = List.Transform(_CharList, each 
            if (List.Contains(_NormalChar,_)) then _ else ";"),
          _DistinctSubj = List.Distinct( Text.Split(Text.Proper( Text.Combine(_ReplaceSep)),";"))
           in List.RemoveItems(_DistinctSubj, {""})
      } }),
    Expand = Table.ExpandListColumn(CleanSubjectsList, "Subjects"),
    Group = Table.Group(Expand, {"Subjects"}, {"Count", each Table.RowCount(_), Int64.Type}),
    Sort = Table.Sort(Group, {{"Subjects", Order.Ascending}})
in
    Sort
