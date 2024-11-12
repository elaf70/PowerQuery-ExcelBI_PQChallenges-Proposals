/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7131853604109385729-RFet

PQ proposal in 3 main steps : 
   1)Prepare 2)Transform 3)Combine

With a dynamic cluster size for groups of {Attribute,Value} + List function usages only
(as always prefer to try to avoid Pivot/Unpivot for complete table)
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData132"]}[Content],
    Prepare = let 
        Chg_Date = Table.TransformColumns(Source,{"Date of Birth", 
            each Date.ToText(Date.From(_), [Format="dd/MM/yyyy"])}),
        Add_FName = Table.CombineColumns(Chg_Date,{"First Name", "Last Name"},
            Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Full Name")
        in Add_FName,
    Transform = let
        Cluster = 2, 
        fxCCN = (size)=>List.Accumulate({1..size}, {}, 
                (s,c)=>s & List.Transform({"Attribute","Value"}, each _&Text.From(c)) ),
        CID = List.Buffer(Prepare[Emp ID]),
        CV = Table.ToColumns(Prepare),
        N = Table.RowCount(Prepare), 
        CA = List.Transform(Table.ColumnNames(Prepare), each List.Repeat({_},N)),
        CAV = List.Combine(List.Skip(List.Zip({CA, CV}))),
        T= List.Accumulate(List.Split(CAV, Cluster*2), {}, 
            (s,c)=> s & {Table.FromColumns({CID} & c, {"Emp ID"} & fxCCN(List.Count(c)/2))}) 
        in T,
    Result = Table.Combine(Transform)
in
    Result


