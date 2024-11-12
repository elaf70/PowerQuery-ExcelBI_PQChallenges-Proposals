/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7235858918315356160-Pvkr

PQ proposal in 4 main steps
    (1) Sort, 
    (2) Group 
    (3) List.Accumulate to generate columns values and names 
    (4) Convert to 1 global table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData214"]}[Content],
    Sort = Table.Sort(Source,{{"Zoo", Order.Ascending}, {"Animals", Order.Ascending}}),
    Group = Table.Group(Sort, "Zoo", {"All", each Table.FromColumns({_[Animals],_[Count]}, {_[Zoo]{0}, _[Zoo]{0}&" Count"}) }),
    Accumulate = List.Accumulate(Group[All], [cols={}, cnames={}], 
        (s,c)=>[cols=s[cols] & Table.ToColumns(c), cnames= s[cnames] & Table.ColumnNames(c)]),
    Result = Table.FromColumns(Accumulate[cols], Accumulate[cnames])
in
    Result



