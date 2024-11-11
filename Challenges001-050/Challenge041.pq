/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7017339031906611200-SPqP
PQ proposal in 6 main steps
1) ...
...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData041"]}[Content],
    FillSeq = Table.FillDown(Source,{"Seq"}),
    Group = Table.Group(FillSeq, {"Seq"}, {"Data", each let 
      _T_1SC = Table.FromRecords({_{0}[[State], [City]]}),
      _TData = let
        _T = Table.PromoteHeaders(Table.RemoveFirstN(_,1)),
        _ToDate = Table.TransformColumns(_T, {{"Date", each Date.From(_)}}),
        _Rename = Table.RenameColumns(_ToDate, {{"Seq", "No"}})
        in _Rename,
      _T_1ST = Table.FromRecords({[Date="Sub Total", Amount=_{0}[Sub Total]]}),  
      _Combine = Table.Combine ({_T_1SC, _TData, _T_1ST}),
      _Fill_SC = Table.RemoveFirstN(Table.FillDown(_Combine,{"State", "City"}),1)
      in _Fill_SC  }),
    Expand = Table.ExpandTableColumn(Group, "Data", {"State", "City", "No", "Class", "Date", "Amount"}),
    GT = List.Sum(Table.SelectRows(Expand, each [Date] = "Sub Total")[Amount]),
    AddRow_GT = Table.Combine({Expand, Table.FromRecords({[Date="Grand Total", Amount=GT]})})
in
    AddRow_GT

