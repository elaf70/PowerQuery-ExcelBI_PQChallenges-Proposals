/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7014802316759347200-vAOE
PQ proposal in 7 main steps
1) ...
...
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData039"]}[Content],
    ChangeType= Table.TransformColumnTypes(Source, {{"Name", type text}, {"Month-Year", type text}, {"Sales", Int64.Type}}),
    Add_DateBOM = Table.AddColumn(ChangeType, "Date", each Date.From([#"Month-Year"]), type date),
    Add_FYear = Table.AddColumn(Add_DateBOM, "FinYear", each Date.Year([Date]+#duration(275,0,0,0)), Int64.Type),

    // GroupBy(Name+FYear) and also be sure of SortOrder for related values before RT calculations
    GroupBy = Table.Group(Add_FYear, {"Name", "FinYear"}, 
        {{"Values", each Table.SelectColumns(Table.Sort(_, "Date"), {"Month-Year", "Sales"}) }}),   
    Transform_Values = Table.TransformColumns(GroupBy,
        { "Values", each
            let
                BufferValues=List.Buffer(_[Sales]),
                ListRT=List.Generate( ()=>[RT=BufferValues{0}, i=0],
                    each [i] < List.Count(BufferValues),
                    each [RT = [RT] + BufferValues{[i]+1}, i=[i]+1],
                    each [RT]
                ),
                ValuesWithRT = Table.FromColumns( Table.ToColumns(_)&{ListRT}, Table.ColumnNames(_)&{"RT"})
            in ValuesWithRT
        }),
    Expand_Values = Table.ExpandTableColumn(Transform_Values, "Values", 
        {"Month-Year", "Sales", "RT"}, {"Month-Year", "Sales", "Running Total"}),

    ChangeType2 = Table.TransformColumnTypes(Expand_Values,{{"Month-Year", type text}, {"Sales", Int64.Type}, {"Running Total", Int64.Type}})
in
    ChangeType2

