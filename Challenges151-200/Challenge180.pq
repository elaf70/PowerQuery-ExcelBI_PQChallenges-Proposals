/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7192734760387960832-6Zh_

PQ proposal in 3 main steps
    1) Add new col with EmployeeName and FillDown
    2) Group by [Emp] to make all calculations in List of RowValues
    3) Back to Table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData180"]}[Content],
    AddEmp = Table.FillDown(Table.AddColumn(Source, "Emp", each if [Sales]=null then [#"Emp-Month"] else null), {"Emp"}),
    Group = Table.Group(AddEmp, "Emp", {"G", each let
        M = List.Skip([#"Emp-Month"]), S=List.Skip([Sales]),
        C = List.Transform(List.Zip({List.Skip(S), List.RemoveLastN(S)}), each Number.Abs(_{1}-_{0})),
        MC = List.Max(C),
        P = List.PositionOf(C, MC) 
        in {[Emp]{0}, List.Sum(S), MC, M{P} & " - " & M{P+1}}
        }),
    Result = Table.FromRows(Group[G], {"Emp","Total Sales","Max Sales CHange","From - To Months"})
in
    Result



