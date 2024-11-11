/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7045605285561786368-rfBh
PQ proposal in 3 main steps
    1) List.Alternate + Combine to get newTable of Names+Values 
    2) Group 
    3) Sort
    + Handling Dynamic ColumNames instead of Hardcoded ones
==================================*/

// SOLUTION Originale : LIST.ALTERNATE+LISTCOMBINE
//===========================================
let
    Source = Excel.CurrentWorkbook(){[Name="tData064"]}[Content],
    ToRows = List.Transform(Table.ToRows(Source), each List.Skip(_)),
    sColNames=List.First(Record.ToList(Source{0})), 
    sColValues=List.First(Record.ToList(Source{1})),
    ToTable = Table.FromColumns(
       {List.RemoveNulls(List.Combine(List.Alternate(ToRows,1,1,1))),      
        List.RemoveNulls(List.Combine(List.Alternate(ToRows,1,1,0))) }, {sColNames, sColValues}),
    Group = Table.Group(ToTable, {sColNames}, {{sColValues, each List.Sum(Table.Column(_, sColValues))}}),
    Sort = Table.Sort(Group,{{sColNames, Order.Ascending}})
in
    Sort


/*
// SOLUTION 2 : TABLE.SPLIT + TRANSPOSE
//===========================================
let
    Source = Excel.CurrentWorkbook(){[Name="tData064"]}[Content],
    ToTable = Table.FromValue(List.Transform(Table.Split(Source,2), 
        each let _NewTable = Table.PromoteHeaders(Table.Transpose(_))
        in Table.SelectRows(_NewTable, each ([Quantity] <> null))) ),
    Expand = Table.ExpandTableColumn(ToTable, "Value", {"Birds", "Quantity"}),
    Group = Table.Group(Expand, {"Birds"}, {{"Quantity", each List.Sum([Quantity])}}),
    Sort = Table.Sort(Group,{{"Birds", Order.Ascending}})
in
    Sort

*/