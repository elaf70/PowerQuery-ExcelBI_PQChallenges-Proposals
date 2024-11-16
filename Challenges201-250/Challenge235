/*================================
Ref = https://www.linkedin.com/feed/update/urn%3Ali%3Aactivity%3A7263400970531848192

PQ proposal in 3 main steps 
    0) Prepare the complete global new column names
    1) Unpivot to get Atrributes and values for {Pop/Male/Female}
    2) Group by Country + inner Group by YearQuarter to combine sub-tables from the unpivoted columns 
    3) Expand inner Tables
==================================*/
let
   Source = Excel.CurrentWorkbook(){[Name="tData235"]}[Content],
   YQ= List.Sort(List.Distinct(Source[#"Year-Quarter"])),
   CN = List.Accumulate(YQ,{}, (s,c)=>s&{c,"Value"&Text.From(List.PositionOf(YQ,c)+1)}),
   Unpivot = Table.UnpivotOtherColumns(Source, {"Country", "Year-Quarter"}, "A", "V"),
   Group = Table.Group(Unpivot, "Country", {"G", each let
     _GroupQ = Table.Group(_, "Year-Quarter", {"Cols", each {[A],[V]}}),
     in Table.FromColumns(List.Combine(_GroupQ[Cols]), CN) }),
   Expand = Table.ExpandTableColumn(Group, "G", CN)
in
 Expand 
