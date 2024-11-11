/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7065536619990138880-WFw5
PQ proposal in 4 main steps
    (1)=Get all values only 
    (2)Transform as RunningTotal 
    (3) Rearrange by subset of 3 with preceedind date 
    (4) Restore as table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData079"]}[Content],
    AllValues = List.Combine(Table.ToRows( Table.RemoveColumns(Source, "Date"))),
    RT = List.Accumulate(AllValues, [T=0, RT={}], (s,c)=>[T=s[T]+c, RT=s[RT]&{s[T]+c}])[RT],
    ToRows = List.Transform(List.Zip({Source[Date],List.Split(RT,3)}), each List.InsertRange(_{1},0,{_{0}})),
    ToTable = Table.FromRows(ToRows, Table.ColumnNames(Source))
in
    ToTable


