/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6974574054083059712-eCGB
PQ proposal in 2 main steps
1) ...
2) ...
==================================*/

let
    Source = Excel.CurrentWorkbook(){[Name="tData008"]}[Content],
    Transform = List.Transform(Source[Data], each let
        L1 = List.RemoveItems(Text.SplitAny(Text.Upper(_), " :,-"),{""}),
        L2 = List.Accumulate(L1, {}, (s,c)=>s & Splitter.SplitTextByCharacterTransition({"A".."Z"}, {"0".."9"})(c)),
        L3 = List.Transform(L2, Text.Proper),
        C = L3{0},
        LR = List.Transform(List.Split(List.Skip(L3),6), each 
           Record.FromList({C}&List.Alternate(_,1,1), {"Company"}&List.Alternate(_,1,1,1)))
        in LR  ),
    Result = Table.FromRecords(List.Combine(Transform), {"Company", "Year", "Revenue", "Cost"})
in 
    Result

