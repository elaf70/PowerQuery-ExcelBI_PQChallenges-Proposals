/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7189835669022621696-9uif

PQ proposal in too-much steps ... 😭

NB : Didn't find on my own the Null replacements or in a very complicated way...😟 
      => had to look for the help of CO-Pilot to get this 1 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData177"]}[Content],
    MCols = List.Skip(Table.ColumnNames(Source),3),
    MergeMarks = let 
      C = Table.CombineColumns(Source, MCols, 
        each [#"Total Marks"=List.Sum(_), Result=if (List.Min(_)<40) then "Fail" else "Pass"], "C")
      in Table.ExpandRecordColumn(C, "C", {"Total Marks", "Result"}),
    AvgMperN = let 
      TG = Table.Group(MergeMarks, "Name", {"AvgM", each List.Average([Total Marks])})
      in Table.Sort(TG, {"AvgM", Order.Descending}), 
    AddRank = Table.AddColumn(MergeMarks, "Rank", each List.PositionOf(AvgMperN[Name], [Name])+1),
    GroupNC = Table.Group(AddRank, {"Name", "Classs"}, {"G", each let
      AddIdx = Table.AddIndexColumn(_,"I", 0),
      R = Table.ReplaceValue(AddIdx, each [I], null, (rv,tv,cv)=> if tv=0 then rv else cv ,{"Name","Classs","Rank"})
      in Table.RemoveColumns(R, "I")}),
    Result = Table.Combine(GroupNC[G])
in
    Result


