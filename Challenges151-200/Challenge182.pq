/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7195271480773926912-6K3j

PQ proposal in 2 main steps
    1) Get Sorted list of [Data] distinct values
    2) With 2 inner Grouping (1st) by [Date] and (2nd) by [Name] with Pivot
        + Date Header values added first for each Group by Date
    => get all expected row values 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData182"]}[Content],
    CNData = List.Sort(List.Distinct(Source[Data])),
    Group = Table.Group(Source, "Date", {"G", each let
        Group = Table.Group(_, "Name", {"G", each let
            TR = Table.ToRows(Table.Pivot(_, CNData, "Data", "Value"))
            in List.Transform(TR, List.Skip) }),
        Header = {Date.ToText(Date.From([Date]{0}), "MM/dd/yyyy" )} & CNData 
        in {Header} & List.Combine(Group[G]) }),
    Result = Table.FromRows(List.Combine(Group[G]))
in
    Result


