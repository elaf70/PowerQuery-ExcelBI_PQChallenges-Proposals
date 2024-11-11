/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-6994509665258598400-50V_
PQ proposal in 3 main steps
1) ...
2) ...
3) ... 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData023"]}[Content],
    ChangeType = Table.TransformColumnTypes(Source,{{"Date", type date}, {"Sales", Currency.Type}}),
    Transform = List.Accumulate( Table.ToRecords(ChangeType), [R={}], (s,c)=>let
       _NbItems = List.Count(s[R]), 
       _IsQMonth = (Number.Mod(Date.Month(_D),3)=0),  
       _D = c[Date],
       _Q = "Q"  & Text.From(Date.QuarterOfYear(_D)),
       _MN = Text.Start(Date.MonthName(_D),3),
       _S = c[Sales],
       _QS = if _IsQMonth then List.Sum(Table.FromRecords(List.LastN(s[R],2))[S])+_S else null,
       _DSPY = if (_NbItems>=12) then _S - s[R]{_NbItems-12}[S] else null,
       _DQSPY = if (_NbItems>=12 and _IsQMonth) then _QS - s[R]{_NbItems-12}[QS] else null,
       _R = s[R] & {[Y=Date.Year(_D), Q=_Q, MN=_MN, S=_S, QS=_QS, DSPY=_DSPY, DQSPY=_DQSPY]}
       in [R=_R] 
    )[R],
    ToTable = Table.FromRecords(Transform),
    Result = Table.RenameColumns(ToTable, { {"Y","Year"}, {"Q","Quarter"}, {"MN","Month Name"}, 
        {"S", "Sales"}, {"QS", "Quaterly Sales"}, {"DSPY", "Yearly Month on Month Sales Diff"}, 
        {"DQSPY", "Yearly Quarter on Quarter Sales Diff"}} )
in
    Result
