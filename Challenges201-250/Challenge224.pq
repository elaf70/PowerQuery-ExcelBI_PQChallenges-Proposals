/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7248542494555070465-Khgb

PQ proposal in 3 main steps
    1) Group by Colum1 with local order and splitting on date value only, and for each Unpivot Data columns
    2) Add [Date] column on each Group inner-tables
    3) Combine all and reorder columns to have [Date] at first 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData224"]}[Content],
    _RenCols = List.Zip({Table.ColumnNames(Source), {"Name"} & List.Skip(Record.FieldValues(Source{0})) }),
    Group = Table.Group(Source, "Column1", 
        {"G", each let
           _d = Date.From(_{0}[Column1]),
           _t = Table.RenameColumns(Table.Skip(_), _RenCols),
           _AddDate = Table.AddColumn(_t, "Date", each _d),
           _Unpivot = Table.UnpivotOtherColumns(_AddDate, {"Date", "Name"}, "Data", "Value") 
           in Table.ReorderColumns(_Unpivot, {"Date", "Name"}) }, 
        GroupKind.Local, (p,c)=>Number.From(try Date.From(c) otherwise 0)  ),        
    Result = Table.Combine(Group[G])
in
    Result


