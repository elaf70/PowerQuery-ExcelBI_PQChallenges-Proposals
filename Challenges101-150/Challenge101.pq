/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7093439911776706561-Uh8R

PQ proposal in 2 main steps
    1) Transform with List.Generate to create all new rows records
    2) Table.FromRecords
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData101"]}[Content],
    Transform = List.Generate(
        ()=>[i=0, rt=0,  nr=[Code=0]],
        each [i]<=Table.RowCount(Source),
        each let 
           _r = Source{[i]},
           _IsNewSB = try (_r[[State],[Branch]]<>[nr][[State],[Branch]]) otherwise true,
           _RT = [rt] + _r[Amount],
           _IsRTMax = (_RT>150)
           in [i  = [i]+1, 
               rt = if (_IsNewSB or _IsRTMax) then _r[Amount] else _RT, 
               nr = _r & [Code=[nr][Code] + Number.From(_IsNewSB or _IsRTMax)] ],
        each [nr] 
    ),
    ToTable = Table.FromRecords(List.Skip(Transform))
in
    ToTable












