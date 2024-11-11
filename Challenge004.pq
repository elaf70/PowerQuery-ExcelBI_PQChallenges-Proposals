/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6970690033053757440-5Nxu
PQ proposal in 5 main steps
1) Order Month short Name
2) Transform ...
3) Group by [Name] to create for each in an new [All] column : 1 record with amounts for all Month + Total 
4) Expand previous Group-calculated column [All] 
5) Sort by [Name] 
==================================*/

let
    Source = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name="tData004"]}[Content],
        {{"Name", type text}, {"Date", type date}, {"Amount", Currency.Type}, {"Order No.", type text}}),
    DCN = List.Distinct(List.Transform(List.Sort(Source[Date]), each Date.ToText(_, [Format="MMM"]))), 
    Transform = let
        T1 = Table.TransformColumns(Source, {
            {"Name", each List.Transform(Splitter.SplitTextByDelimiter("/")(_), each Text.Trim(_))},
            {"Date", each Date.ToText(_, [Format="MMM"])} }),
        T2 = Table.ReplaceValue(T1, each [Amount],
            each Number.Round([Amount]/List.Count([Name]),2), Replacer.ReplaceValue,{"Amount"}),
        T3 = Table.ExpandListColumn(T2, "Name")
        in T3,
    Group = Table.Group(Transform, {"Name"}, {"All", each let
        _O = Text.Combine([#"Order No."],","),
        _P = Table.Pivot(_[[Date],[Amount]], 
            List.Distinct(List.Transform([Date], Text.From)), "Date", "Amount", List.Sum),
        _T = List.Sum([Amount])
        in Record.Combine( {[#"Order Nos"=_O], _P{0}, [Total=_T]})  }),
    Expand = Table.ExpandRecordColumn(Group, "All", {"Order Nos"}& DCN & {"Total"} ),
    Sort = Table.Sort(Expand, {"Name", Order.Ascending})
in
    Sort
