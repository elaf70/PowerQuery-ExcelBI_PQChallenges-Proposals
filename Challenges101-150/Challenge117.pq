/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7113734211152338944-ITFd

PQ proposal in 2 main steps
    1. Add column with 1 new table from List.Transform 
        to get 4 relatives date formated + payment part
    2. Project to only [Company + new added column] and expand it 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData117"]}[Content],
    Add = Table.AddColumn(Source, "Custom", each let
        _D = [Date], _P = [Payment]/4,  
        _V = List.Transform({0..3}, each let 
            _d = Date.AddQuarters(_D,_),
            _sQ = "Q" & Text.From(Date.QuarterOfYear(_d)),
            _sY =  DateTime.ToText(_d, "-yy") 
            in {_sQ & _sY, _P} )
        in Table.FromRows (_V,{"Quarter", "Payment"}) ),
    Expand = Table.ExpandTableColumn(Add[[Company],[Custom]],
        "Custom", {"Quarter", "Payment"})
in
    Expand


