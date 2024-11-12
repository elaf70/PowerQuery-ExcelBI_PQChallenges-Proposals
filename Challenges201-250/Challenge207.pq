/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7227888316916895744-G7Zp

PQ proposal in 4 main steps
    1) Replace indicator values "Y" by related Name
    2) Create inner-tables with All-names per day using List.Accumulate
    3) Combine all these inner-tables as final result 

As usual : I try also to use only List functions to avoid more obvious 1st idea of Unpivot/Group/... that in most case don't scale well if needed 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData207"]}[Content],
    DayNames = List.Skip(Table.ColumnNames(Source)),
    ReplaceName = Table.ReplaceValue(Source, "Y", each [Name], Replacer.ReplaceValue, DayNames),
    Transform = List.Accumulate(DayNames, {}, (s,c)=>let
        _l = List.RemoveNulls(Table.Column(ReplaceName, c)),
        _cn = {"Day of Week"} & List.Transform({1..List.Count(_l)}, each "Name"&Text.From(_))
        in s &  {Table.FromRows({{c} & _l}, _cn)} ),
    Combine = Table.Combine(Transform)
in
    Combine




