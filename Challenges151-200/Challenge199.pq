/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7217739351147872256-3VjL

PQ proposal in 3 main steps
    1) Split each string and get for each 1 table of NumPart+Date
    2) Combine subtables + change Date type (as not kept in Combine !)
    3) Sort 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData199"]}[Content],
    CN = {"Part No.", "Date"},
    Split = List.Transform(Source[String], each let
        L = Text.SplitAny (Text.Replace(_,"//","/"), " ,"),
        pNum = List.Select(L, each try Number.From(_)<>null otherwise false){0},
        Dates= List.Select(L, each try Date.From(_)<>null otherwise false)
        in Table.FromRows(List.Transform(Dates, each {pNum, _}), CN) ),
    T = Table.TransformColumnTypes(Table.Combine(Split), {"Date", type date}),
    Sort = Table.Sort(T, List.Zip({CN, {0,0}}))
in
    Sort



