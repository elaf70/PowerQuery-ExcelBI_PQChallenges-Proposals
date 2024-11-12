/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7215565192485621760-otuB

PQ proposal in 4 main steps
    1) Add a Month column (=detail values)
    2) GroupByMonth to find then Max value in each
    3) Buidi The RunningTotal of all these Groups
    4) Join Detail value + RTByMonth with next 1 projection of needed cols 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData198"]}[Content],
    Group = Table.Group(Source, {"Date"}, 
        {{"T", each _}, {"Max", each List.Max([Value])}}, 
        GroupKind.Local, (x,y)=>let f=each Date.StartOfMonth(_[Date]) in Value.Compare(f(x),f(y)) ),
    RT = List.Accumulate(Group[Max], {}, (s,c)=>s & {(List.Last(s)??0)+ c} ),
    Add_RT = Table.FromColumns({Group[T], RT}, {"T", "Running Total"}),
    Expand = Table.ExpandTableColumn(Add_RT, "T", Table.ColumnNames(Source))
in 
    Expand


