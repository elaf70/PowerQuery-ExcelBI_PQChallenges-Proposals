/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7202881638760955904-dSDC

PQ proposal in 2 main steps
    1) Tranform rows
    2) Combine
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData188"]}[Content],
    Transform = List.Transform(Table.ToRows(Source), each let
        _Store = _{0}, _DStart = Date.From(_{1}), _DEnd = Date.From(_{2}), 
        _Days = List.Dates(_DStart, Number.From(_DEnd-_DStart)+1, #duration(1,0,0,0)),
        _DailyAmount = _{3}/List.Count(_Days), 
        _T = List.Transform(_Days, each let
            _Q = Text.Format("Q#{0}-#{1}", {Date.QuarterOfYear(_),Text.End(Text.From(Date.Year(_)),2)})
            in {_Store, _Q}),
        _Group = Table.Group(Table.FromRows(_T, {"S", "Q"}), "Q", 
            {"G", each {[S]{0},[Q]{0}, Number.Round(Table.RowCount(_)*_DailyAmount)}} ) 
        in Table.FromRows(_Group[G], {"Store","Quarter","Amount"}) ),
    Result = Table.Combine(Transform)
in
    Result

