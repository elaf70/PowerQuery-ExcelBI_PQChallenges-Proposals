/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7050308778574983168-iA57
PQ proposal in 4 main steps
    1) Group by [Data1] with GroupKindLOcal and change as detected a new "Name"  for each Transpose & PromoteHeader
    2) Combine result and Filldows [House] 
    3) Filter
    4) Sort
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData067"]}[Content],
    Group = Table.Group(Source,"Data1", {"T", each Table.PromoteHeaders(Table.Transpose(_))},
      GroupKind.Local, (b,e)=>Number.From(List.Contains({"House No.","Name"},e))),
    Combine = Table.Combine(Group[T]),
    FillDown = Table.FillDown(Table.RenameColumns(Combine,{"House No.", "House"}),{"House"}),
    FilterRows = Table.SelectRows(FillDown, each ([Name] <> null)),
    Sort = Table.Sort(FilterRows,{{"House", Order.Ascending}, {"Name", Order.Ascending}})
in
    Sort


