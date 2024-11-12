/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7139101359873716224-MtwR

PQ proposal in 3 main steps
    1) set Table2 as 1 Record with prices of each company
    2) Transform [Company] column of Table1 into inner tables by spliting text , retrieving Price from Record of step1 and building each table fromRecords
    3) Expand Table column of step2
==================================*/
let
    fxSource = (n)=>Excel.CurrentWorkbook(){[Name=n]}[Content],
    T2 = fxSource("tData137_2"),
    RCs = Record.FromList(T2[Price], T2[Company] ),

    T1 = fxSource("tData137_1"),
    CN= {"ID","Company","Price"},
    Transform = Table.TransformColumns(T1, {"Company", each let
        _S = Text.SplitAny(Text.Remove(_," "),";,"),
        _R = List.Transform(_S, each let 
            s = Text.Split(_,":")
            in Record.FromList(s & {Record.Field(RCs,s{1})}, CN) ),        
        _T = Table.FromRecords(_R) 
        in Table.Sort(_T,{"Company",Order.Ascending}) }),
    Expand = Table.ExpandTableColumn(Transform, "Company", CN)
in
    Expand


