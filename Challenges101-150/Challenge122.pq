/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7119170026447732737-mMxG

PQ proposal in 3 main steps
    1) Split [DateTime] column
    2) Group by [Vehicule]+[Date] to create inner table for each with related records
    3) Combine
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData122"]}[Content],
    Split_DT = Table.SplitColumn(Source, "Date Time", each let
        _DT = DateTime.From(_) 
        in {DateTime.Date(_DT),DateTime.Time(_DT)}, {"Date", "Time"}),
    Group = Table.Group(Split_DT, {"Vehicle", "Date"}, {"T", each let
        _R1 = _{0}[[Vehicle],[Date]],
        _R2 = List.Accumulate(List.Zip({{1..Table.RowCount(_)},[Time]}), [],
            (s,c) =>  Record.AddField(s, "Time"&Text.From(c{0}), c{1}) ),
        _D = List.Accumulate(Table.ToRecords(_), 1, 
             (s,c)=>s + (if (c[Status]="Stop") then 1 else -1)*Number.From(c[Time])),
        _R3 = [Total=Time.From(Number.Mod(_D,1))]
        in Table.FromRecords({Record.Combine({_R1,_R2,_R3})})
        }),   
    Combine = Table.Combine(Group[T])
in
    Combine


