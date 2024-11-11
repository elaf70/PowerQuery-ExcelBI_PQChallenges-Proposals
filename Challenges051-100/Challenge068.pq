/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7050678718565949440-CMI1
PQ proposal in 3 main steps
    1) Group & Expand
    2) Sort
    3) Add Row for Totals 
==================================*/
let
    _MMM = List.Transform({1..12}, each Date.ToText(#date(1900, _,1),[Format="MMM", Culture="en-US"])),

    Source = Excel.CurrentWorkbook(){[Name="tData068"]}[Content],
    Group =  Table.Group(Source, "Company", {
        {"Sales", each List.Sum([Sales])},
        {"MMM", each let 
            _D = List.Combine(Table.TransformRows(_, each {Number.From([From])..Number.From([To])})),  
            _M = List.Transform(_D, each let 
                _D = Date.From(_)
                in if (Date.DayOfWeek(_D, Day.Monday)<5) then Date.ToText(_D, [Format="MMM", Culture="en-US"]) else null ),
            _T = Table.FromColumns({List.RemoveNulls(_M)},{"M"}),
            _G = Table.Group(_T, "M", {"C", List.Count}),
            _P = Table.Pivot(_G, _MMM, "M", "C", List.Sum)
            in  Table.AddColumn(_P, "Row Total", each List.Sum(Record.FieldValues(_)))} 
        }),
    Expand = Table.ExpandTableColumn(Group, "MMM", _MMM & {"Row Total"}),
    Sort = Table.Sort(Expand,{{"Company", Order.Ascending}}),
    AddRow_CTotal = let
        _CurCols = Table.ToColumns(Table.DemoteHeaders(Sort)),
        _NewCols = List.Transform(_CurCols, each _ & {if (_{0}="Company") then "Column Total" else List.Sum(List.Skip(_))} )
        in Table.PromoteHeaders(Table.FromColumns(_NewCols))
in
    AddRow_CTotal


