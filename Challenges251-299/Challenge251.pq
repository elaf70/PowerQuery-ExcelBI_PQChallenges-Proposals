/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7283694165534593024-qabx

PQ proposal in 3 main steps
  (1) Group by EmpID with unpivot to create 1 inner table for each
  (2) Combine all of these
  (3) Split column Name => First + Last 

==================================*/

let
    Source = Excel.CurrentWorkbook(){[Name="tData251"]}[Content],
    Group = Table.Group(Source, {"Emp ID"}, {"G",(t)=>let
        _L = Table.UnpivotOtherColumns(t, {"Emp ID"}, "A", "V")[V],
        _Rec = Record.FromList(List.Alternate(_L,1,1), List.Alternate(_L,1,1,1))
        in Table.FromRecords({Record.SelectFields(t{0}, {"Emp ID"}) & _Rec}) }),
    Combine = Table.Combine(Group[G]),
    SplitName = Table.SplitColumn(Combine, "Full Name", Splitter.SplitTextByEachDelimiter({" "}), {"First Name", "Last Name"})
in
    SplitName


