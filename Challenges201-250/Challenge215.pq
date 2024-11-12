/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7238033260621033472-tXtm

PQ proposal in 4 main steps
    1) Change Date type
    2) Filter Rows with no paiement already
    3) GroupBy Customer to build for each subset expected new rows values using List.Accumulate
    4) Combine all to create final-table 
==================================*/
let
    Today = Date.From(DateTime.LocalNow()),
    Source = Excel.CurrentWorkbook(){[Name="tData215"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Due Date", type date}, {"Paid Date", type date}}),
    Filter = Table.SelectRows(ChangeType, each [Paid Date]=null),
    Group = Table.Group(Table.Sort(Filter, {"Branch ID", 0}), "Branch ID", 
        {"G", each let
            _Sort = Table.Sort(_, {{"Customer", 0},{"Due Date", 0}}),
            _FirstRow = {_Sort[Branch ID]{0},null,null,null},
            _Acc = List.Accumulate(Table.ToRecords(_Sort), {{_Sort[Branch ID]{0},null,null,null}}, (s,c)=>let
                _Cust = if (c[Customer]=List.Last(s){0}) then null else c[Customer]
                in s & {{_Cust, c[Due Date], c[Loan Amt], -Duration.Days(c[Due Date]-Today)}} )
        in _Acc },
        GroupKind.Local),
    ToTable = Table.FromRows(List.Combine(Group[G]), 
        {"Branch ID /Customer","Due Date","Total Loan Amount", "Total Outsanding Days"})
in
    ToTable



