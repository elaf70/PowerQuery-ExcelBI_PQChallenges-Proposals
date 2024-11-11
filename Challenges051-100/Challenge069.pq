/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7050678718565949440-CMI1
PQ proposal in 2 main steps
    1) Group by [Machine] to create for each a Table for each Date+Start
    2) Expand
==================================*/
let
    RCN = {"Status", "Start Time", "End Time"},
    Source = Excel.CurrentWorkbook(){[Name="tData069"]}[Content],
    Group = Table.Group(Table.RenameColumns(Source,{{"Start Date Time", "Start Time"}}), {"Machine"}, 
        {"D", each let
            _Sort = Table.Sort(_, {"Start Time", Order.Ascending}),
            _Rec = Table.ToRecords(_Sort),
            _Transform = List.Accumulate(List.Skip(_Rec), {_Rec{0}}, (s,c)=>let
                _Last = List.Last(s),
                _NewStatus = c[Status]<>_Last[Status]
                in if _NewStatus then List.RemoveLastN(s) & {Record.AddField(_Last, "End Time", c[Start Time]), c} else s),
            _ToTable = Table.FromRecords(List.RemoveLastN(_Transform,1))
            in try Table.SelectColumns(_ToTable, RCN) 
                otherwise Table.FromRows(List.Repeat({null}, List.Count(RCN)), RCN)  
        }),
    Expand = Table.ExpandTableColumn(Group, "D", RCN)
in
    Expand


