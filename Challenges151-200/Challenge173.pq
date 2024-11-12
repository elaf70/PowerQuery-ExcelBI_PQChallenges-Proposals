/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7184762226812043264-C841

PQ proposal in 5 main steps
    ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData173"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Add_ListYQM = Table.AddColumn(ChangeType, "L", 
        each {Date.Year([Date]), Date.QuarterOfYear([Date]), Date.ToText([Date], "MMM", "en-US")}),
    GroupYQM = Table.Group(Add_ListYQM, {"L"}, {{"Y", each [L]{0}{0}},{"S", each List.Sum([Sale])}}),
    GroupY = Table.Group(GroupYQM, "Y", {"G", each let
        TYQM = List.Accumulate(List.Skip([L]), [r={[L]{0}}, pq=[L]{0}{1}], 
            (s,c)=>[r=s[r] & {{null,if(c{1}=s[pq]) then null else c{1},c{2}}},pq=c{1}] )[r],
        SSY = List.Sum([S]), 
        TS = List.Transform([S], each {_, _/SSY}),
        TYRow = {Text.Combine({Text.From([Y]{0})," Total"}), null, null, SSY, 1},  
        Rows = List.Transform(List.Zip({TYQM, TS}), List.Combine)  & {TYRow}
        in Table.FromRows(Rows, {"Year", "Quarter","Month","Total Sales", "Sales %"}) 
        }),
    Combine = Table.Combine(GroupY[G])
in
    Combine

