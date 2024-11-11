/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-...
PQ proposal in 3 main steps
1) Group by [Name] to create for each a record for all WeekdayName with proper indicator "Y" 
2) Expand record from previous Grouped column with proper DayName ordering 
3) Sort by [Name]
==================================*/

let
    Source = Excel.CurrentWorkbook(){[Name="tData007"]}[Content],
    Group = Table.Group(Source, {"Name"}, 
        {"G", each List.Accumulate(List.RemoveNulls([Day of Week]), [], (s,c)=>Record.AddField(s,c,"Y"))}),
    Expand = Table.ExpandRecordColumn(Group, "G", {"Sun", "Mon", "Tue","Wed","Thu","Fri","Sat"} ),
    Sort = Table.Sort(Expand, {"Name", Order.Ascending})
in 
    Sort
