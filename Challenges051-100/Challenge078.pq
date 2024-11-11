/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7063362290758090752-quPl
PQ proposal 
    Using code with 1 "try" & MULTIPLE "otherwise" to calculate age values from the first sucessul of the 3 kind of "Lookup" to perform
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData078_1"]}[Content],
    T2 = Excel.CurrentWorkbook(){[Name="tData078_2"]}[Content],
    Add_FullName = Table.CombineColumns(T2, {"First Name", "Last Name"}, Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"FullName"),
    Mail1 = List.Buffer(T2[E Mail 1]),
    Mail2 = List.Buffer(T2[E Mail 2]),
    FullName = List.Buffer(Add_FullName[FullName]),
    Age = List.Buffer(Add_FullName[Age]),
    Add_Age = Table.AddColumn(Source, "Age", each 
        try Age{List.PositionOf(Mail1, [E Mail])} 
        otherwise try Age{List.PositionOf(Mail2, [E Mail])} 
        otherwise try Age{List.PositionOf(FullName, [Name])} 
        otherwise null  )
in
    Add_Age


