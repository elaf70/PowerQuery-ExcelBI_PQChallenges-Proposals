/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7179688805430398976-XrX3

PQ proposal in 1 main steps with Text.Split + List functions

NB : But still missing native RegularExpression available in M 😭... 😉 
     This would have been so far easier & simplier to handle then 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData169"]}[Content],
    Result = Table.AddColumn(Source,"Codes", each let
        WL2 = List.Select(Text.SplitAny([String]," ,"), each Text.Length(_)>=2),// Words having at least 2 char
        W = List.Select(WL2, each let                                           // Keeping only those that 
           C = Text.ToList(_)
           in List.Contains({"A".."Z"}, C{0})                                   // - Start with a Letter
                and List.ContainsAny(C,{"0".."9"})                              // - Contains Numbers  
                and List.IsEmpty(List.RemoveItems(C, {"A".."Z"}&{"0".."9"}))    // - With no other char than Upper-Alphabet + Numbers 
            )
        in Text.Combine(W, ", ")
        )
in
    Result


