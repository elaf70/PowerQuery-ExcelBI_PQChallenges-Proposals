/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7265937111965442048-0T9b

PQ proposal in 3 main steps
1) Get col-Names
2) Transform each cols to replace Null as required with List.Accumulate
3) Rebuild with it, the final Result table
==================================*/
let
 Source = Excel.CurrentWorkbook(){[Name="tData237"]}[Content],
 CNames = Table.ColumnNames(Source),
 Transform = List.Transform(List.Zip({CNames, Table.ToColumns(Source)}), each 
    List.Accumulate(_{1}, [n=_{0}, i=0, r={}], (s,c)=>[n=s[n], i=s[i]+Number.From(c=null), r=s[r] & {c??(n&"_"&Text.From(i))}])[r] ),
 Result = Table.FromColumns(Transform, CNames) 
in
 Result
