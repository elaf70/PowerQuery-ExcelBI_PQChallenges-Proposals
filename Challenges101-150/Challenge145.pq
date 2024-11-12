/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7149254806904229888-sblQ

PQ proposal in 3 main steps
    1) Calculate an initial Record of RunningTotal parameters for all stores
        For Each store ={CurrentRT value : 0 at start, FirstDate }
    2) Transform each RowRecord of Source table with List.Accumulate 
        to add new RT field 
        Using & updating the Record of parameters each time as requested
    3) final result = Table.FromRecords
==================================*/
let
  Source = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name="tData145"]}[Content],
    {{"Date", type date}}),
  RecRTP = let 
    _G = Table.Group(Source, {"Store"}, {"DMin", each Date.From(List.Min([Date]))})
    in Record.FromList(List.Transform(_G[DMin], each {0,_}), _G[Store]),
  Transform = List.Accumulate(Table.ToRecords(Source),[RP=RecRTP, r={}], (s,c)=>let
       _SP = Record.Field(s[RP], c[Store]),
       _DD12M = c[Date]<Date.AddMonths(_SP{1},12),
       _NewSP = if (_DD12M) then {_SP{0}+c[Sale], _SP{1}} else {c[Sale], c[Date]},
       _NewRP = Record.TransformFields(s[RP],{c[Store], each _NewSP}) 
       in [RP=_NewRP, r=s[r]&{Record.AddField(c, "Column1", _NewSP{0})}] ),
  Result = Table.FromRecords(Transform[r])
in
  Result


