/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7142000487175819264-8ypm

PQ proposal in 5 main steps
    1) Set Initial stock from Table2 as record
    2) Calculate List of all EndOFDays stock values from Table1
    3) Calculate List of indicator(1/0) if Last entry of group [Date,Item]
    4) Calculate List of all new EODS replacing null for which are not last
    5) Add column with new EODS
==================================*/
let
    TStock = Excel.CurrentWorkbook(){[Name="tData140_2"]}[Content],
    RStock = Record.FromList(TStock[Stock], TStock[Item]),
    Source = Excel.CurrentWorkbook(){[Name="tData140_1"]}[Content],
    CType = Table.TransformColumnTypes(Source,{{"Date", type date}, {"Time", type time}}),
    EODS = List.Accumulate(Table.ToRows(CType), [S=RStock, R={}], (s,c)=>let
        _NewS = Record.Field(s[S], c{2}) - c{3} 
        in [S = Record.TransformFields(s[S], {c{2}, each _NewS}),
            R = s[R] & {_NewS}] )[R],
    IsLast = List.Combine(Table.Group(Source, {"Date", "Item"}, 
        {"IsLast", (t)=>List.Repeat({0}, Table.RowCount(t)-1) & {1}})[IsLast] ),
    NewEODS = List.Accumulate(List.Zip({EODS, IsLast}), {}, (s,c)=> s & {if (c{1}=1) then c{0} else null}),
    Result = Table.FromColumns(Table.ToColumns(CType) & {NewEODS}, Table.ColumnNames(CType)&{"EODS"})
in
    Result



