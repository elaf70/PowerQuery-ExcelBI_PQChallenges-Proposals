/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7076045877273931776-Kt9Q
PQ proposal in 2 main steps
    1) Transform with ListAccumulate to add new indexes to each RowRecord + keeping 1 'copy' of current Indexes values at each step
    2) Reset TransformedRecords-List to final Table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tDataO88"]}[Content],
    Transform = List.Accumulate(Table.ToRecords(Source), [R={}, Idx={0,0,0}], (s,c)=>let
        _NewIdx = if (c[Country]<>null) then {s[Idx]{0}+1, 1, 1}
              else if (c[State]<>null) then {s[Idx]{0}, s[Idx]{1}+1, 1}
              else {s[Idx]{0}, s[Idx]{1}, s[Idx]{2}+1},
        _NewR = c & [#"Index Country"=_NewIdx{0}, #"Index State"=_NewIdx{1}, #"Index City"=_NewIdx{2}]      
        in [R = s[R] & {_NewR}, Idx= _NewIdx])[R],
    ToTable = Table.FromRecords(Transform)    
in
    ToTable


