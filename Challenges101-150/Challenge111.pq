/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7106124064921915392-HWhj

PQ proposal in 2 main steps
    1. Add colum with a Record with all requested new fields calculation.
       Mainly from List transformation 
    2. Expand this column 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData111"]}[Content],
    Add_Col = Table.AddColumn(Source, "Custom", each let 
        _M = [Teams Members],
        _LP = List.Accumulate(Text.Split([Individuals], ", "), {}, (s,c)=>let 
                _P=Text.PositionOf(_M,c) 
                in s & {[Ind=c,Pos=if (_P=-1) then null else _P+1, Len=Text.Length(c)]} ),
      _T = Table.Sort(Table.FromRecords(_LP), {"Len", Order.Descending}),
      _R = List.Accumulate(Table.ToRecords(_T), [], (s,c)=> let 
                _I="Ind"&Text.From(Record.FieldCount(s)+1)              
                in  s & Record.FromList({c[Ind], c[Pos]}, {_I, _I & "_Pos"}) )
      in _R),
    Expand = Table.ExpandRecordColumn(Add_Col, "Custom",
        List.Union(List.Transform(Add_Col[Custom], each Record.FieldNames(_))))
in
    Expand










