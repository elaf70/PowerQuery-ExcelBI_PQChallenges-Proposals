/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7068073334239899648-OGMc
PQ proposal in 2 main steps
    1) Transform to generate for each Subject the TOP3 student & Note
    2) Create Result-Table from columns
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData081"]}[Content],
    Columns = Table.ToColumns(Source),
    Names = Columns{0},
    Subjects = List.Skip(Table.ColumnNames(Source)),

    Transform = List.Transform(List.Skip(Columns), each let
        _Notes = _,
        _Max3 =  List.MaxN(List.Distinct(_Notes),3),
        _Sel = List.Select(List.Zip({Names, _Notes}), each List.Contains(_Max3,_{1})),
        _Table = Table.FromRows(_Sel, {"Name","Note"} ), 
        _Grp = Table.Group(_Table, "Note", {"GN", each Text.Combine(_[Name], "#(lf)")})
        in Table.Sort(_Grp, {"Note", Order.Descending})[GN] ),
    ToTable = Table.FromColumns( {{1,2,3}} & Transform, {"Rank"} & Subjects )
in
    ToTable


