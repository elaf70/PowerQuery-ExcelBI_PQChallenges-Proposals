/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7207592773540057088-3q61

PQ proposal in 2 main steps
    1) Transform
    2) To Table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData191"]}[Content],
    fxSplit = Splitter.SplitTextByAnyDelimiter,
    Digits = {"0".."9"}, Alphabets={"A".."Z"}, 
    Transform = List.Transform(Source[Text], each let
       _Words = List.Select(Text.Split(_, " "), each Text.Upper(_)=_),
       _Split = List.Transform(_Words, each let 
            w = Text.Select(_, Alphabets & Digits)
            in  List.Select(fxSplit(Digits)(w) & fxSplit(Alphabets)(w), each _<>"")),
       _Select = List.Select(_Split, each List.Count(_)=2),
       _Combine = Text.Combine(List.Transform(_Select, each Text.Combine(_,"_")),", ")
        in _Combine),
    Result = Table.FromColumns({Transform}, {"Answer Expected"})
in
    Result
.

