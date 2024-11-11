/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7045247499187412992-OefV
PQ proposal in 11 main steps
    1) ...
    ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData063"]}[Content],
    ListDates = List.Skip(Record.ToList(Source{0}),2),
    RenCols = Table.RenameColumns(Source, 
        {{"Column1","Product"}} & List.Zip({List.Skip(Table.ColumnNames(Source),2),ListDates})),
    Add_Plant = Table.AddColumn(RenCols, "Plant", each if [Column2]=null then [Product] else null),
    FillDown = Table.FillDown(Add_Plant, {"Product", "Plant"}),
    RemoveNull = Table.SelectRows(FillDown, each ([Column2] <> null)),
    Add_Type = Table.AddColumn(RemoveNull, "Type", each if([Product]=[Plant]) then "Plant" else "Product"),
    Unpivot = Table.UnpivotOtherColumns(Add_Type, {"Plant", "Type", "Product", "Column2"}, "Date", "Value"),
    Group = Table.Group(Unpivot, {"Plant","Date"}, {{"Data", each let
        _t = Table.Group(Table.RemoveColumns(_,{"Plant","Date"}), {"Type"},
            {"Data", each Table.Pivot(_, List.Distinct(_[Column2]), "Column2", "Value") }),
        _pl= Table.RemoveColumns(_t{List.PositionOf(_t[Type],"Plant")}[Data], {"Type","Product"}),
        _pr= Table.RemoveColumns(_t{List.PositionOf(_t[Type],"Product")}[Data], {"Type"}),
        _ac = Table.AddColumn(_pl, "P", each _pr)
        in Table.ExpandTableColumn(_ac, "P", Table.ColumnNames(_ac[P]{0})) }}),
    Expand = Table.ExpandTableColumn(Group, "Data", Table.ColumnNames(Group[Data]{0})),
    ChangeType = Table.TransformColumnTypes(Expand,{{"Date", type date}}, "en-GB"),
    Sort = Table.Sort(ChangeType,{{"Date",0}, {"Plant",0}, {"Product",0}})
in Sort


