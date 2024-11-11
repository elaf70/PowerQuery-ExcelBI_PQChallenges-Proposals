/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7060825576223252480-QcYf
PQ proposal in 4 main steps
    (1)=GroupBy "Country" + Transform to table for each 
    (2)=Expand Result-tables
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData076"]}[Content],
    YQ_Names = List.Alternate(Table.ColumnNames(Source),1,1),
    Group = Table.Group(Source, {"Country"}, {"All", each let
       _Cols=List.Skip(Table.ToColumns(_)),
       _Zip=List.Zip({List.Alternate(_Cols,1,1,1), List.Alternate(_Cols,1,1), YQ_Names}),
       _ToRecord = List.Transform(_Zip, each Record.AddField(Record.FromList(_{1}, _{0}), "Year-Quarter", _{2}))  
       in Table.FromRecords(_ToRecord)
    }),
    Expand = Table.ExpandTableColumn(Group, "All", {"Year-Quarter", "Population", "Male", "Female"})
in
    Expand

