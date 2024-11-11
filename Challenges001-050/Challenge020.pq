/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6989797395572109312-e6hp
PQ proposal in 3 main steps
1. Get FieldNames 
2. GroupBy EmpID with for each 
 - Get the 2 rows of data as 2 Records 
 - Select FieldNames with Distinct Value in these 2 Records 
 - Combine SelectedNames + return it as list x2 as required.
3. Expand 2 lines per Group 
==================================*/

let
    Source = Excel.CurrentWorkbook(){[Name="tData20"]}[Content],
    FieldNames = List.Skip(Table.ColumnNames(Source),2),
    Group = Table.Group(Source, {"Emp ID"},
      {{"Mismatched Columns", 
        each let 
            // Get the 2 Records of each GROUP of EmpID
            R = Table.ToRecords(_),
            // Select FieldNames with Distinct Value in the 2 Records  
            FWDV = List.Select(FieldNames, (f)=>Record.Field(R{0}, f) <> Record.Field(R{1}, f) ),
            // Combine their names
            CDF = Text.Combine(FWDV, ", ")
        in  {CDF, CDF} }} ),
    Expand = Table.ExpandListColumn(Group, "Mismatched Columns")
in Expand
