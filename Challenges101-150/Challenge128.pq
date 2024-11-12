/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7126780171843424256-Tljm

PQ proposal in 2 main steps
    1) Prepare the list of all SoundCode for each Letter
    2) Add new column to source table with following substeps 
    -Get Letters from each LastName excluding 1st & those to disregard

==================================*/
let
    T2 = Excel.CurrentWorkbook(){[Name="tData128_2"]}[Content],
    LS = List.Buffer(List.Accumulate(Table.ToRows(T2),{}, (s,c)=> let 
        _T = List.Transform(Splitter.SplitTextByDelimiter(",")(c{1}), each {Text.Trim(_), Text.From(c{0})})
        in s & _T)),

    Source = Excel.CurrentWorkbook(){[Name="tData128"]}[Content],
    Add_SC = Table.AddColumn(Source, "Soundex Code", each let
        _S = Text.Upper(List.Last(Text.Split([Name], " "))), 
        _LC =  List.Select(List.Skip(Text.ToList(_S)), each Text.PositionOf("AEIOUHWY", _)=-1),
        _RemDupl = List.Accumulate(_LC, {}, (s,c)=>if (c=List.Last(s)) then s else s&{c}),
        _CS = Text.Combine(List.ReplaceMatchingItems(_RemDupl, LS)),
        _C = Text.Start(_S,1) & "-" & Text.PadEnd(Text.Start(_CS,3), 3, "0")
        in _C)
in
    Add_SC


