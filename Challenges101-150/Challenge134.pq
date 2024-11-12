/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7134390317272711168-QwJJ

PQ proposal in 2 main steps
    1) Group by Store with LIst.Accumulate to calculate & add [ClosingStock]
    2) Table.Combine as global result
==================================*/
let
    t1 = Excel.CurrentWorkbook(){[Name="tData134"]}[Content],
    t2 = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData134_2"]}[Content]),
    Group = Table.Group(t1, {"Store"}, {"All", (t)=> let
        _OpenStk = t2[Open Stock]{List.PositionOf(t2[Store],t[Store]{0})},
        _Transform = List.Accumulate(Table.ToRecords(t), [Stk=_OpenStk, r={}],  (s,c)=> let
            _NewStk = s[Stk]+c[IN]-c[OUT]
            in [Stk=_NewStk, r=s[r]&{Record.AddField(c, "Closing Stock", _NewStk)}])
        in Table.FromRecords(_Transform[r]) }),
    Combine = Table.Combine(Group[All])
in
    Combine

