/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7200344905100009472--bbY

PQ proposal in 2 main steps
    1) Merge by taking firt row of the combination of delivery items on same day or those from previous/next day
    2) Expand 
==================================*/
let
    D = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData186b"]}[Content]),
    fxSelectDelivery = (dates)=>Table.SelectRows(D, each List.Contains(dates, [Delivery Date])), 
    Source = Excel.CurrentWorkbook(){[Name="tData186"]}[Content],
    Merge = Table.AddColumn(Source, "M", each let  
        d = [Calendar Date],
        s = fxSelectDelivery({d}) & fxSelectDelivery({Date.AddDays(d,-1),Date.AddDays(d,+1)}) 
        in Table.FirstN(s,1) ),
    Expand = Table.ExpandTableColumn(Merge, "M", Table.ColumnNames(D))
in
    Expand


