/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6979649401036124160-1-dO
PQ proposal in 3 main steps
1) Sort
2) Group(ID) & Transform each subset by adding a new column with Sequence List from current Seq to Seq on next line
3) Expand 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData012"]}[Content],
    Sort = Table.Sort(Source,{{"ID", Order.Ascending}, {"Seq", Order.Ascending}}),

    // Group(ID) & Transform each subset 
    //   By Adding a new column with Sequence List from current Seq to Seq on next line 
    //   and then expand this new List column to add all new records   
    Group =Table.Group(Sort, {"ID"}, {{"Data", each let
        _Add_Index = Table.AddIndexColumn(_, "Index", 0,1),
        _Add_ListSeq = Table.AddColumn(_Add_Index, "ListSeq", 
             each {[Seq]..([Seq]+(try _Add_Index[Seq]{[Index]+1} - [Seq] -1 otherwise 0))}),
        _DelCol = Table.RemoveColumns(_Add_ListSeq, {"Seq", "Index"})
        in Table.ExpandListColumn(_DelCol, "ListSeq")
    }}),
    Expand = Table.ExpandTableColumn(Group[[Data]], "Data", 
        {"ID", "Data1", "Data2", "ListSeq"}, {"ID", "Data1", "Data2", "Seq"})
in
    Expand
