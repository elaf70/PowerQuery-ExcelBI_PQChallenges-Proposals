/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7157220754260996096-cNOQ

PQ proposal in 2 main steps
    1) Group by [Name] to generate all Dates of each period removing WE then Group [Name,Date] to calculate Nb of consecutive days of each
    2) Combine all previous result table
==================================*/
let
    OrderType = {"ML", "PL", "CL"},
    Source = Excel.CurrentWorkbook(){[Name="tData152"]}[Content],
    Group = Table.Group(Source, {"Name"}, {"All", each let 
        Add_Order = Table.AddColumn(_, "TypeOrder", each List.PositionOf(OrderType, [Type of Leave])),
        CombineDate = Table.CombineColumns(Add_Order,{"From Date", "To Date"},
            each {Int64.From(Date.From(_{0}))..Int64.From(Date.From(_{1}))},"Date"),
        ExpandDate = Table.ExpandListColumn(CombineDate, "Date"),
        SelectWD = Table.SelectRows(ExpandDate, each Date.DayOfWeek(Date.From([Date]),Day.Monday)<5),
        Group = Table.Group(SelectWD, {"Name", "Date"}, {"Type", each let
            _T = if (Table.RowCount(_)=1) then _ else Table.Sort(_,{"TypeOrder", Order.Ascending})  
            in _T[Type of Leave]{0} }),
        Pivot = Table.Pivot(Group, OrderType, "Type", "Date", List.Count)
        in Pivot }),
    Result = Table.Combine(Group[All])
in
    Result

