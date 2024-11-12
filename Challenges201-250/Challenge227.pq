/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7253253537017585664-f7gC

PQ proposal in 4 main steps
    1) Add in T2 : 2 match-function from their rule definitions
    2) Add in T1 Salary col. with result of match-function of T2
    3+4) as match-results can be multiple/null : expand & removeNull

NB =A little "tricky" but allows to scan T1 only once. 
==================================*/
let
    Source = (t)=>Excel.CurrentWorkbook(){[Name=t]}[Content],
    T2AddFns = Table.Buffer(Table.FromRecords(
        Table.TransformRows(Source("tData227_2"), each let
            _Fns = List.Transform({[Sequence], [Name]}, each 
                if (_="*") then (s)=>true
                else if Text.EndsWith(_,"*") then (s)=>Text.StartsWith(s,Text.BeforeDelimiter(_,"*")) 
                else if Text.StartsWith(_,"*") then (s)=>Text.EndsWith(s,Text.AfterDelimiter(_,"*"))
                else (s)=>s=_)
            in _ & [FnSeq=_Fns{0}, FnName=_Fns{1}]) 
        )),    
    T1AddMatchSalary = Table.AddColumn(Source("tData227_1"), "Salary", each let
         _S = Text.From([Sequence]), _N=Text.Upper(Text.From([Name])), 
        _Result = Table.SelectRows(T2AddFns, each [FnSeq](_S) and [FnName](_N))[Salary]
        in _Result),
    Expand = Table.ExpandListColumn(T1AddMatchSalary, "Salary"),
    RemoveNull = Table.SelectRows(Expand, each ([Salary] <> null))
in
    RemoveNull


