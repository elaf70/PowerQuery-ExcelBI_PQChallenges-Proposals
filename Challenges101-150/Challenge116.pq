/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7111559882373251072-LnWI

PQ proposal in 2 main steps
    1) Transform List of row-values by adding Forecasts with ListGenerate 
    2) Back for table, providing also new Forecast ColumnNames 
==================================*/
let
    //---- Round Number to Nearest M-multiple : exple fxMRound(13,5)=15 
    fxMRound = (n,m)=>Number.Round(n/m)*m,     

    Source = Excel.CurrentWorkbook(){[Name="tData116"]}[Content],
    Transform = List.Transform(Table.ToRows(Source), each let
        _PM = List.Skip(_), _NM = List.Count(_PM),
        _F = List.Generate( ()=>[i=0, f=fxMRound(List.Sum(_PM)/_NM, 10)], 
            each [i]<12-_NM,  
            each [i=[i]+1, f=fxMRound([f]*1.05, 5)],
            each [f] )
        in _ & _F),
    ToTable = Table.FromRows(Transform, 
        Table.ColumnNames(Source) & {"Jul","Aug","Sep","Oct","Nov","Dec"})
in
    ToTable


