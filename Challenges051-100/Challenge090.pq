/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7078582576923222016-2D4_
PQ proposal in 2 main step :  
    1) Transform with List.Accumulate to generate list of row values
    2) Back to result table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData090"]}[Content],
    CN=Table.ColumnNames(Source),
    Transform = List.Accumulate (Table.ToRows(Source), [r={}, i=0, n=1], (s,c)=>let
        _CheckNewSerie = Number.From((s[i]+1)=s[n]),
        _NewCN = if(_CheckNewSerie=1) then {CN} else {}
        in [i = (s[i]+1)*(1-_CheckNewSerie), n = s[n]+_CheckNewSerie, r = s[r]&{c}&_NewCN] ),
    ToTable = Table.FromRows(Transform[r], CN)
in
    ToTable


