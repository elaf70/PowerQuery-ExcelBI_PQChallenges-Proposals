/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7085830345538846720-RscZ
PQ proposal in 3 main steps
    (1) Reorder Mark items 
    (2) Transform Data in 1 RecordsList 
    (3) Table.FromRecords 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData095"]}[Content],
    ReOrderMark = Table.SelectColumns(Source,
       List.RemoveItems(Table.ColumnNames(Source),{"Marks2"}) & {"Marks2"}),
    Transform= List.Accumulate(Table.ToRecords(ReOrderMark), {}, (s,c)=>let 
        _L = List.Select(List.Zip({Record.FieldNames(c),Record.ToList(c) }), (l)=>(l{1}<>null)),
        _NbRec= List.Count(_L) / 2,
        _Records = List.Accumulate({0.._NbRec-1}, {}, 
            (s,c) => s & {Record.FromList({_L{c}{1},_L{c}{0},_L{c+_NbRec}{1}}, {"Class","Subject","Mark"})} )
        in s & _Records),
    ToTable = Table.FromRecords(Transform)
in
    ToTable


