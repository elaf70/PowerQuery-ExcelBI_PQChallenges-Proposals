/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7210129241378639872-vLfS

PQ proposal in 6 main steps
    ...

For fun try not to use obvious Transpose function ... 😉
    Because as always : not scale well if needed
    Not so easy, but quite sure there should also be sth simplier (?) 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData193"]}[Content],
    Categs = List.Distinct(List.Skip(Record.FieldValues(Source{1}))),
    NCateg = List.Count(Categs),
    Persons = List.Skip(Source[Column1],2),
    SplitByQ = Table.TransformRows(Table.Skip(Source,2), each let
        _VAll = Record.FieldValues(_),
        _P = List.First(_VAll)
        in  List.Transform(List.Split(List.Skip(_VAll), NCateg), each _ &{List.Sum(_)})),
    Accumulate = List.Accumulate(List.Skip(SplitByQ), [r=SplitByQ{0},p=SplitByQ{0}], (s,c)=>let
        _New=List.Transform(List.Zip({s[p],c}), each List.Transform(List.Zip(_), List.Sum))
        in [r=List.Transform(List.Zip({s[r],_New}), List.Combine), p=_New])[r],
    PCol = List.Accumulate(Persons, {}, (s,c)=>let
        _New = Text.Combine({List.Last(List.RemoveNulls(s)),c},", ")
        in s & {_New} & List.Repeat({null},NCateg) ),
    CCol = List.Repeat(Categs & {"Total"}, List.Count(Persons)),
    Result = Table.FromColumns({PCol,CCol} & Accumulate, {"Person","Category","Q1","Q2","Q3","Q4"})
in
    Result


