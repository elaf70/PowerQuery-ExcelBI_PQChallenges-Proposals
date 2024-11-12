/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7187661342286237696-tA5J

PQ proposal in 3 main steps
    1) Define a function that expand 2 strings as List of items & values 
        into a table with all items and RunningTotalOfValues in columns
    2) Combine Column1 & 2 with this function
    3) Expand subTables as expected result
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData176"]}[Content],
    //--- Function : To Expand 2 string as List of items &  values (separator = ", ")
    //     into  a table with all items and RunningTotalOfValues in columns
    fxExpandToTable = (sListItems as text, sListValues as text)=>let
        I = Text.Split(sListItems, ", "), V= Text.Split(sListValues, ", "),
        A = List.Accumulate({0..List.Count(I)-1}, [r={}, rt=0], (s,c)=>let 
            _rt=s[rt]+(try Number.From(V{c}) otherwise 0) 
            in [r=s[r] & {{I{c},_rt}}, rt=_rt]) 
        in Table.FromRows(A[r]) ,
    CN = {"Column1", "Column2"},
    CombineC = Table.CombineColumns(Source, CN, each fxExpandToTable(_{0}, _{1}), "C"), 
    Expand = Table.ExpandTableColumn(CombineC, "C", CN)
in
    Expand


