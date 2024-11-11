/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7043068570695516160-SKWA
PQ proposal in 5 main steps
    1) ...
    ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData062"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Date", type date}}),
    Group = Table.Group(ChangeType, {"Emp ID"}, {{"Data", each let
        _Dates = List.Select(_[Date], each List.Contains({0,6}, Date.DayOfWeek(_))=false),
        ToRecFT = List.Generate(()=>[x=1, y=[Fr=_Dates{0},To=_Dates{0}]],
            each [x]<=List.Count(_Dates),
            each let 
                _d=_Dates{[x]}, _Dur = Duration.Days(_d-[y][To]),
                _Fr = if (_Dur=1 or (Date.DayOfWeek(_d, Day.Monday)=0 and _Dur=3)) then [y][Fr] else _d
                in [x=[x]+1, y=[Fr=_Fr,To=_d]],
            each [y] ),
        GrpDate = Table.Group(Table.FromRecords(ToRecFT), {"Fr"}, {{"To",each List.Max([To])}, {"WD", each Table.RowCount(_)}})
        in GrpDate }}),
    ReplaceError = Table.ReplaceErrorValues(Group, {{"Data", Table.FromRecords({[Fr="",To="",WD=0]})}}),
    Expand = Table.ExpandTableColumn(ReplaceError, "Data", {"Fr","To","WD"}, {"From Date", "To Date", "Total WorkDays"}),
    ToText = Table.TransformColumns(Expand,{
        {"From Date", each try Date.ToText(_,"dd/MM/yyyy") otherwise "NA"},
        {"To Date"  , each try Date.ToText(_,"dd/MM/yyyy") otherwise "NA"} })
in
    ToText


