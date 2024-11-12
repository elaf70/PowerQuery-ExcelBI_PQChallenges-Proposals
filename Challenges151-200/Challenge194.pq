/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7210490795810869248-_kbm

PQ proposal in 3 main steps
    1 + 2) Get 2 List of RunningTotalValues and Previous (with 0 at first) 
    3) Calculate OriginalValueList : each = RT - Previous
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData194"]}[Content],
    CNames = Table.ColumnNames(Source), 
    RTValues = List.Combine(Table.ToRows(Table.RemoveColumns(Source,{"Date"}))),
    Previous = {0} & List.RemoveLastN(RTValues,1),
    Values = List.Transform(List.Zip({RTValues,Previous}), each _{0}-_{1}),
    Split = List.Split(Values, List.Count(CNames)-1),
    Rows = List.Transform(List.Zip({Source[Date], Split}), each {_{0}} & _{1}),
    ToTable = Table.FromRows(Rows, CNames)
in
    ToTable


/*
// ---- SOLUTION 2 : Same logic but row/row with ListGenerate =>  which should scale better if needed

let
    Source = Excel.CurrentWorkbook(){[Name="tData194"]}[Content],
    Rows = Table.ToRows(Source), NRows=List.Count(Rows),
    Transform = List.Generate(()=>[r=1, p=0, s=Rows{0}], each [r]<=NRows, 
        each let  
            _RV = Rows{[r]},
            _V = List.Skip(_RV), 
            _P = {[p]} & List.RemoveLastN(_V,1), 
            _T = List.Transform(List.Zip({_V,_P}), each _{0}-_{1})
            in [r=[r]+1, p=List.Last(_RV), s={_RV{0}} & _T],
        each  [s]),
    Result = Table.FromRows(Transform, Table.ColumnNames(Source))
in

*/

