/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7101413016759377920-GV0f

PQ proposal in 2 main steps
    1) ...
    2) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData108"]}[Content],
    No = List.Zip(
        List.Accumulate(Source[Serial],{},(a,v)=> let 
            n = Text.Length(Text.Select(v,".")), 
            b = List.Last(a,{}),
            y = List.Last(b) = Text.BeforeDelimiter(v,".",n-1), 
            c = if n = 0 then {v} else List.FirstN(b,n) &{v} 
            in List.RemoveLastN(a, Number.From(y)) & {c}) ),
    Ans = Table.FromColumns({No{0}} & List.Transform(No,each List.Transform(_,each Source{[Serial=_]}?[Names]?)),
          {"Serial"}& List.Transform({1..List.Count(No)},each "Name"&Text.From(_)) )
in
    Ans

