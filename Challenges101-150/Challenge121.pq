/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7118807636094914560-8LB6

PQ proposal in 4 main steps
    1) Add Group column 
    2) Fill down [Group] 
    3) Remove Null [Name] 
    4) GroupBy [Group] with FillUp & Down of [Dept] values
==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData121"]}[Content],
  Add_G = Table.AddColumn(Source, "Group", 
      each if Text.StartsWith(Text.From([Emp ID]), "Group") then [Emp ID] else null), 
  FillDown_G = Table.FillDown(Add_G, {"Group"}),
  Remove_NullN = Table.SelectRows(FillDown_G, each [Name]<>null),
  Transform_G = Table.Combine(
     Table.Group(Remove_NullN, {"Group"}, 
        {"T", each Table.FillUp(Table.FillDown(_, {"Dept"}), {"Dept"})})[T] )  
in
  Transform_G











