/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7035096037992087552-6U83
PQ proposal with following main steps
    1. Add 2 Columns Hall&Date + FillDow & Filter 
    2. Group by Hall+Date with a Table.Pivot for GuestList data
    3. Expand resullting inner-Table for GuestData 
 ==================================*/

// 1ere SOLUTION publiée : 
let
    Source = Excel.CurrentWorkbook(){[Name="tData055"]}[Content],
    Add_Hall = Table.AddColumn(Source, "Hall", each if ([Data1]="Hall") then [Data2] else null),
    Add_Date = Table.AddColumn(Add_Hall, "Date", each if ([Data1]="Date") then [Data2] else null),
    FillDown = Table.FillDown(Add_Date,{"Hall", "Date"}),
    Filter = Table.SelectRows(FillDown, each ([Data1] <> "Date" and [Data1] <> "Hall")),
    AllGuests= List.Distinct(Filter[Data1]),
    Group = Table.Group(Filter, {"Hall","Date"}, {{"Data",        each Table.Pivot(_[[Data1],[Data2]], List.Distinct(_[Data1]), "Data1", "Data2")}}),
    Expand = Table.ExpandTableColumn(Group, "Data", AllGuests, AllGuests)
in
    Expand






/*
// 2nd SOLUTION : With 2 successive Table.Group with GroupKind.local
//==================================================================
let
    Source = Excel.CurrentWorkbook(){[Name="tData055"]}[Content],
    Group1 = Table.Group(Source, {"Data1"}, 
        { {"Record", each Record.FromList([Data2],[Data1])} },
        0,(x,y)=> Number.From(List.Contains({"Hall","Date"}, y[Data1])) ),
    Group2 = Table.Group(Group1, {"Data1"}, 
        { {"Records", each List.Transform(List.Skip([Record]),(f)=>[Record]{0} & f)} }, 
        0, (x,y)=>Number.From(y[Data1]="Hall") ),
    ToRecords = Table.ExpandListColumn(Group2[[Records]], "Records")[Records],
    TableFromRecords = Table.FromRecords(ToRecords, List.Distinct(Source[Data1]),MissingField.UseNull)
in
    TableFromRecords
*/
