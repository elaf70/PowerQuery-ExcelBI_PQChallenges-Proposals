/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7095977203238080512-eyLw

PQ proposal in 4 main steps
    1) Transform with List.Accumulate to create new Records 
    2) To create 1 table with it 
    3) Expand [Team] column
    4) Sort by [Team]
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData103"]}[Content],
    GCN = List.Select(Table.ColumnNames(Source),
        each ( Text.StartsWith(_,"G") and                     Value.FromText(Text.RemoveRange(_,0,1)) is number) ),

    Transform = List.Accumulate(Table.ToRecords(Source), {}, (s,r)=> let 
        _LR = List.Accumulate(GCN, {}, (s,g)=>let
            d = Record.Field(r,g),
            t = Text.Split(Text.From(Record.Field(r,d)), ", ")
            in s & (if (d=null) then {} else {[Group=g, Dept=d, Team=t]}) )
        in s & _LR ),
    ToTable = Table.FromRecords(Transform),
    Expand = Table.ExpandListColumn(ToTable, "Team"),
    Sort = Table.Sort(Table.TransformColumnTypes(Expand,{{"Team", Int64.Type}}),
        {{"Team", Order.Ascending}})
in
    Sort

