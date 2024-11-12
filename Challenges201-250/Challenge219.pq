/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7243106677049176064-RAwO

PQ proposal in 3 main steps
    1) Get List of all possible distinct Devices 
    2) Transform [Machine] column into inner-tables handling possible missing Device/OS
    3) Expand it
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData219"]}[Content],
    Devices = List.Distinct(List.Transform(Source[Machine], each Text.Trim(Text.SplitAny(_, "-,"){0}))),
    Transform = Table.TransformColumns(Source, {"Machine", each let
        _R = List.Transform(Text.Split(_, ", "), each 
            if Text.Contains(_, " - ") 
                then Text.Split(_, " - ") 
                else if List.Contains(Devices, _) then {_,null} else {null,_}),
        _T = Table.FromRows(_R, {"Device", "OS"}),
        _Fill = Table.FillUp(Table.FillDown(_T,{"Device"}),{"OS"})
        in _Fill}),
    Expand = Table.ExpandTableColumn(Transform, "Machine", {"Device", "OS"})
in
    Expand



