/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7035458425287258112-Nuqo
PQ proposal in 3 main steps
    1. Clean Fruit Names : ProperFormat + Remove final 's' (as trim)
    2. Group with 1 List.Accumulate to create 1 Record with all Fruits data in each Garden
    3. Expand All these Garden Records, to get final expected result
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData056"]}[Content],
    Clean = Table.TransformColumns(Source, { {"Fruit", each Text.TrimEnd(Text.Proper(Text.Trim(_)),"s")} }),
    Add_Count = Table.AddColumn(Clean, "Count", each 1),
    Pivot = Table.Pivot(Add_Count, List.Distinct(Add_Count[Fruit]), "Fruit", "Count", 
      each if  (List.Sum(_)=1) then "Unique" else ("Not Unique-"& Number.ToText(List.Sum(_))) ) 
in
    Pivot

