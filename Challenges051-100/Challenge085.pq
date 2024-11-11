/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7073146763527733248-8h2S
PQ proposal in 5 main steps
    (1) Few Preparation stuff 
    (2) Group to sum all Revenue & Cost
    (3) AddRow : GrandTotal 
    (4) AddColumn : TotalProfit 
    (5) Rename Columns as expected
==================================*/
let 
    Source = Excel.CurrentWorkbook(){[Name="tData085"]}[Content],  

    //-- Sum all table values 
    fxTSum=(t as table)=>List.Sum(Table.ToList(t,List.Sum)),             
    //-- Sum all table values from a list of Columns Names only
    fxTCSum=(t as table, cn as list)=>fxTSum(Table.SelectColumns(t, cn)),  
    CN = Table.ColumnNames(Source),
    CN_Rev = List.Select(CN, each Text.Contains(_,"Rev")),
    CN_Cost = List.Select(CN, each Text.Contains(_,"Cost")),
    Fill_Company = Table.FillDown(Source,{"Company"}),

    Group = Table.Group(Fill_Company, {"Company"}, 
        {{"TR", each fxTCSum(_, CN_Rev) },{"TC", each fxTCSum(_, CN_Cost) }} ),

    Append_GT = Table.Combine({Group, 
        Table.FromRecords({[Company="Grand Total", TR=List.Sum(Group[TR]), TC=List.Sum(Group[TC])]})}),
    Add_TP = Table.AddColumn(Append_GT, "TP", each [TR]-[TC]),

    Rename = Table.RenameColumns(Add_TP,{{"TR", "Total Revenue"}, {"TC", "Total Cost"}, {"TP", "Total Profit"}})
in
    Rename


