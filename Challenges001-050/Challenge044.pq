/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7020238135297028096-N0jo
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData044"]}[Content],
    ColNames = List.Skip(Table.ColumnNames(Source)),
    Years = List.Distinct(List.Transform(ColNames, each Text.Middle(_, 0,4))),
    NewColNames = List.Transform(Years, each _ & " Total") & {"Till Date Total"},
    Group = Table.Group(Source, "Item", {"Data", each let
        _SubTable= _,
        _YValues = List.Accumulate(Years, {}, (s,c)=>let
            _CN = List.Select(ColNames, each Text.StartsWith(_,c)),
            _Cols = Table.ToColumns(Table.SelectColumns(_SubTable,_CN)),
            _Values = List.RemoveNulls(List.Combine(_Cols))
            in  s & {List.Sum(_Values)} ),
        _AllValues = _YValues & {List.Sum(_YValues)} 
        in Record.FromList(_AllValues, NewColNames) }),
    Expand = Table.ExpandRecordColumn(Group, "Data", NewColNames),
    Add_TotalRow = let
        _NewCols = List.Transform(Table.ToColumns(Expand), each _ & {try List.Sum(_) otherwise "Total"})
        in  Table.FromColumns(_NewCols, Table.ColumnNames(Expand))
in
    Add_TotalRow

