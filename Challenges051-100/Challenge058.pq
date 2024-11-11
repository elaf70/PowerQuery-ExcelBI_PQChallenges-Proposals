/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7037995142133297152-tinT
PQ proposal in 4 main steps
    1) Add 1 new column with OB-values From 1 List.Accumulate
    2) Group By Day with Table.Min on OB value
    3) Expand record of each day 
    4) Rename column according to specification 
==================================*/
let
    Opening=100,
    Source = Excel.CurrentWorkbook(){[Name="tData058"]}[Content],
    OB = List.Accumulate(List.RemoveLastN(Source[Txn Amount],1), {Opening}, (s,c)=>s & {List.Last(s)+c}),
    Add_OB = Table.FromColumns({Source[Txn Date], Source[Serial], OB}, {"Txn Date","Serial","OB"}),
    Group = Table.Group(Add_OB, {"Txn Date"}, {"Data", each Table.Min(_, "OB")}),
    Expand = Table.FromRecords(Group[Data], {"Txn Date", "Serial", "OB"}),
    Rename = Table.RenameColumns(Expand,{{"OB", "Min OpeningBalance"}}),
    ChangeType = Table.TransformColumnTypes(Rename,{{"Txn Date", type date}})
in
    ChangeType

