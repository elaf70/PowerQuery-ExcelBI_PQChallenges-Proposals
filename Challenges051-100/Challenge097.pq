/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7088367062804291584-RUpY
PQ proposal in 2 main steps
    1) Transform in record values
    2) Back to Result table

Simple Transpose & Filldown would also do the job. 
But this won't be extendable & efficient on "bigger table" if needed. 
=> So I always prefer List functions approach :Even if for this one, code is a little bit longer & "complex"... ;=) 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData097"]}[Content],
    _CN = Table.ColumnNames(Source), 
    TransformRows = Table.TransformRows(Source, each let
      _T = List.Accumulate(Record.ToList(_), [p=null, r={}], (s,c)=>let 
         _New = if (c=null and s[p]<>null) then s[p] else c  
         in [r=s[r]&{_New}, p=_New])
      in Record.FromList(_T[r], _CN) ),
    ToTable = Table.FromRecords(TransformRows)
in
    ToTable


