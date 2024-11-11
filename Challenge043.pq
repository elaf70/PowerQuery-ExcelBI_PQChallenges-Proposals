/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7019875747842433024-kLXG
PQ proposal in 4 main steps
1) ...
2) ...
3) ... 
4) ...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData043"]}[Content],
    CN = Table.ColumnNames(Source),  CNSub = List.Skip(CN,3),
    TransformRecords = List.Accumulate(Table.ToRecords(Source), {}, (s,c)=>let
        _NbSub = List.NonNullCount(Record.FieldValues(c))-3,
        _NewRec = if (_NbSub=1) then {c} else let
                _RPart1 = Record.TransformFields(c[[Name],[Date],[Hours]], {"Hours", (x)=>x/_NbSub}),
                _LR = List.Accumulate(CNSub, {}, 
                    (s1,c1)=>if (Record.Field(c,c1)="X") then s1&{Record.AddField(_RPart1, c1, "X")} else s1) 
                in _LR
        in s & _NewRec),
    ToTable = Table.FromRecords(TransformRecords, CN, MissingField.UseNull),
    ChangeType = Table.TransformColumnTypes(ToTable,{{"Date", type date}})
in
    ChangeType 

