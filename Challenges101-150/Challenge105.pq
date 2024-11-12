/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7098513916334342144-OMiO

PQ proposal in 1 main steps
    1) Add colum Duration calculted from Split and convert in Duration according to each content and a final Sum of all items
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData105"]}[Content],

    fxToDuration = (sValue, sType)=>let
        dTypeConvert = {{"d",#duration(1,0,0,0)}, {"h",#duration(0,1,0,0)}, 
                        {"m",#duration(0,0,1,0)}, {"s",#duration(0,0,0,1)} },
        _Type = Text.Lower(Text.Start(sType,1))
        in try 
            Number.FromText(sValue) * List.Select(dTypeConvert, each _{0}=_Type){0}{1} 
        otherwise #duration(0,0,0,0),
    
    Result = Table.AddColumn(Source, "Duration", each let 
        _S = [Text Time] & " " & [Add],
        _DParts = List.Transform(
            List.Split(Splitter.SplitTextByDelimiter(" ", QuoteStyle.Csv)(_S), 2),
            each fxToDuration(_{0}, _{1}))
        in List.Sum(_DParts) )
in
    Result

