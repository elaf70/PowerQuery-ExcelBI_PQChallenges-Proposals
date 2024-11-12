/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7103949732686815232-fCge

PQ proposal in 2 main steps
    1) ...
    2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData110"]}[Content],
    Transform = Table.TransformRows(Source, each let
      _FN = Record.FieldNames(_),  _FV = Record.FieldValues(_),
      _NV = Text.Split(_FV{0},", "),   _Nb = List.Count(_NV),
      _NbCheck = List.Count(List.Select(_FV, each _="X")),
      _RemainV = Text.Combine(List.LastN(_NV, if (_NbCheck<_Nb) then _Nb-_NbCheck else 0), ", "),
      _CNV =  List.Select(List.Zip({List.Skip(_FN), List.Skip(_FV)}), each _{1}="X"),
      _RDetails = List.Accumulate(_CNV, [], 
        (s,c) => s & Record.FromList({if (c{1}="X") then _NV{Number.Mod(Record.FieldCount(s),_Nb)} else ""}, {c{0}}) )
      in [Names=_FV{0}] & _RDetails & [Remaining=_RemainV] ),
    Result = Table.FromRecords(Transform, Table.ColumnNames(Source) & {"Remaining"}, MissingField.UseNull)
in
    Result


