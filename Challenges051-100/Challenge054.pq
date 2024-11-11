/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7032921710144172032-IVgt
PQ proposal in 6 main steps
    1) ...
    ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData054"]}[Content],
    ChangeType = Table.TransformColumnTypes({{"Date", type date}}),

    Dates = { Number.From(List.Min(Source[Date]))..Number.From(List.Max(Source[Date])) },
    TDateAction = let    
      _T = Table.AddColumn(Table.FromColumns({Dates},{"Date"}), "Action", each {"Buy","Sell"})
      in Table.TransformColumnTypes(Table.ExpandListColumn(_T, "Action"), {{"Date", type date}}),

    LData = List.Accumulate(List.Distinct(Source[Product]), {}, (s,c)=>let
      _T = Table.AddColumn(TDateAction, "Product", each c, type text),
      _J =  Table.NestedJoin(_T, {"Product","Date","Action"}, Source, {"Product","Date","Action"}, "Amt", JoinKind.LeftOuter),
      _E = Table.TransformColumns(_J, {"Amt", each try _[Amt]{0} otherwise 0}),
      _B = List.Accumulate(Table.ToRecords(_E), {0}, (s,c)=>s & {List.Last(s) + c[Amt]*(if(c[Action]="Buy") then 1 else -1)}),
      _Result = Table.FromColumns(Table.ToColumns(_E) & {List.Skip(_B)}, Table.ColumnNames(_E)&{"Balance"}) 
      in s & {_Result}),
    Combine = Table.Combine(LData),
    ReOrder = Table.ReorderColumns(Combine, {"Date","Product","Action"}),
    Sort = Table.Sort(Combine,{{"Date", Order.Ascending}, {"Product", Order.Ascending}, {"Action", Order.Ascending}})
in Sort

