/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6981833403625025536-4Kqt
PQ proposal in 3 main steps
1) Group by [Group] & Transform to create for each with List.Generate 1 table with related Values and RunningTotal   
2) Expand column
3) Expand record
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData013"]}[Content],
    Group_Transform = Table.Group(Source, {"Group"}, {
        {"Value", each let
            _Val = List.Buffer(_[Value]),
            _NbItems = List.Count(_Val),
            _RT = List.Generate(
                ()=>[RT=if(_Val{0}<0) then 0 else _Val{0}, i=0],
                each [i] < _NbItems,
                each [RT=if(_Val{i}<0) then 0 else [RT] + _Val{i}, i=[i]+1],
                each [RT])
            in Table.FromColumns({_Val, _RT}, {"V", "RT"}) }
    } ),
    ExpandColumn = Table.ExpandListColumn(Group_Transform, "Value"),
    ExpandRecord = Table.ExpandRecordColumn(ExpandColumn, "Value", {"V", "RT"}, {"Value", "Running Total"})
in
    ExpandRecord

