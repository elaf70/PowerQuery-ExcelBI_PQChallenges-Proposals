/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7223175352724934657-ODc1

PQ proposal in 3 main steps
    1) Buffer all original columns content, without null values
    2) Transform each with list of related matche(s) in other columns
    3) Build with it the FinalResult table using Table.FromColumns
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData204"]}[Content],
    IndexList = {0..Table.ColumnCount(Source)-1}, 
    Cols = List.Buffer(List.Transform(Table.ToColumns(Source), each List.RemoveItems(_, {null, ""}))),
    NewCols = List.Transform(IndexList, each let
        _Values = Cols{_},
        _NewValues = List.Accumulate(List.RemoveItems(IndexList, {_}), {}, (s,c)=>let
           _NbMatch = List.Count(List.Intersect({_Values,Cols{c}}))
           in s & (if _NbMatch=0 then {} else {"Col"&Text.From(c+1)&" - "&Text.From(_NbMatch)}) )
        in _NewValues),    
    Result = Table.FromColumns(NewCols, List.Transform(Table.ColumnNames(Source), each _ & " Match"))
in
    Result


