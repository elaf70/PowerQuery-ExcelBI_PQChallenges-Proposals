/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7136564640230035456-hBYa

PQ proposal in 3 main steps
    1) Get Postition for the Min of BinSize
    2) Add new colum with a record as the expected "split"
    3) Expand new column with RecordFiedNames of PosMinBinSize 
==================================*/
let
    MaxB = 10,
    Source = Excel.CurrentWorkbook(){[Name="tData135"]}[Content],
    PosMinBS = List.PositionOf(Source[Bin size], List.Min(Source[Bin size])),
    Add_RG = Table.AddColumn(Source, "RG", each let
        b = [Bin size],
        n = Number.IntegerDivide(MaxB,b) + Number.From(Number.Mod(MaxB,b)>0), 
        r = List.Accumulate({1..n},[], (s,c) => Record.AddField(s, "Group " & Text.From(c), 
            "B"&Text.From((c-1)*b+1)&"-B"&Text.From(List.Min({c*b, MaxB}))) ) 
        in r),
    Expand = Table.ExpandRecordColumn(Add_RG, "RG", Record.FieldNames(Add_RG[RG]{PosMinBS}))
in
    Expand


