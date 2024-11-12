/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7197805781231030272-IN3B

PQ proposal in 2 main steps
    1) Group by [Set] and Split each [Text] item, Select last one with expected conditions and Combine matching ones
    2) Add [New Count] column 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData184"]}[Content],
    NChar = {"0".."9"}, LChar = {"A".."Z","a".."z"},
    Group = Table.Group(Source, "Set", {
        {"Text", each let 
            _L = List.Transform([Text], each let
                _Split = Splitter.SplitTextByCharacterTransition(NChar,LChar)(_),
                _Select = List.Select(_Split, each 
                    List.Contains(LChar, Text.Start(_,1)) and List.Contains(NChar, Text.End(_,1)))
                in List.Last(_Select))
            in Text.Combine(_L, "-")
        },
        {"Original Count", List.Count}} ),
    Add_NewC = Table.AddColumn(Group, "New Count", each List.Count(Text.Split([Text], "-")))
in
    Add_NewC


