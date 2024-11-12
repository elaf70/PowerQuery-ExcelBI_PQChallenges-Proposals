/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7151784947173777409-5Jnw

PQ proposal in 2 main steps
    1) List.Accumulate for each original table rows to store progressively
        listofvalues of each column until having all filled
        and if listOfValues completed then add in the Result-List 
        1 innerTable with repeated records + suffix[Type] by added index 
    2) Combine InnerTables from previous step
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData147"]}[Content],
    CN = Table.ColumnNames(Source), NC = List.Count(CN), LVNull = List.Repeat({null}, NC),
    Transform = List.Accumulate(Table.ToRows(Source), [r={}, v=LVNull, i=0], (s,c)=>let
        _V = List.Transform(List.Zip({c, s[v]}), each _{0}??_{1}),
        _VC = List.NonNullCount(_V) = NC,
        _NewT = if (_VC=false) then null else let
            _R = Record.FromList(_V, CN),
            _T = Table.FromRecords(List.Repeat({_R}, s[i]+1)),
            _TI = Table.CombineColumns(Table.AddIndexColumn(_T,"_Idx_",1), 
                {"Type", "_Idx_"}, each _{0}&Text.From(_{1}), "Type")
            in _TI
        in if (_VC) then [r=s[r] & {_NewT}, v=LVNull, i=0] else [r=s[r], v=_V, i=s[i]+1] ),
    Result = Table.Combine(Transform[r])
in
    Result


