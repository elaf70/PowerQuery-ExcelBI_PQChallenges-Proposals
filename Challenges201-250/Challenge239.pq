/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7268473832686977024-mLNf

PQ proposal in 3 main steps
1) Get all Ranking Value disctinct & Sorted : for each rows = [Q4]-[Q1]
2) Transform original rows in new row-records with 2 List.Accumulate
3) final result = Table.FromRecord with MissingField.UseNull

==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData239"]}[Content],
  RankVal = List.Buffer(List.Sort(List.Distinct(Table.TransformRows(Source, each [Q4]-[Q1])), Order.Descending)),
  TransformRows = List.Accumulate(Table.ToRows(Source), {}, (s,r)=>let
    _SplitRow = List.Accumulate({1..3}, {}, (s,c)=>let 
      _Rec = [#"QtQ Drop" = "Q"&Text.From(c+1)&"-Q"&Text.From(c), Amount=r{c+1}-r{c}] 
      in s & {_Rec & (if (c=1) then [Name=r{0},Rank=List.PositionOf(RankVal, r{4}-r{1})+1] else [])}) 
    in s & _SplitRow),
  Result = Table.FromRecords(TransformRows, {"Name","QtQ Drop","Amount","Rank"}, MissingField.UseNull)
in
  Result
