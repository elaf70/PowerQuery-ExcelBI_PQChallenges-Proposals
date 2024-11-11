/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7002480754316185600-yWUO
PQ proposal in 1 main steps 
1) Use Table.TransformColumns to perform all transformations needed to split items and then recombine it as  required
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData030"]}[Content],
    Transform = Table.TransformColumns(Source, {"Highest Goal Scorer", each let
        _RD = Text.FromBinary(Text.ToBinary(_, 1251),  TextEncoding.Ascii), // Replace Diacritics
        _RX = Text.Remove(_RD, Text.ToList("()")),                          // Remove xTra char
        _SN = Text.Remove(_RX, {"0".."9"}),                                 // NotNumeric Char 
        _SplitNum = List.RemoveItems(Text.SplitAny(_RX,_SN),{""}),
        _SplitTxt = Splitter.SplitTextByCharacterTransition({"a".."z"},{"A".."Z"})(_SN)
        in Text.Combine({_SplitNum{0},_SplitTxt{0},_SplitTxt{1},_SplitNum{1}},"-") })
in
    Transform

