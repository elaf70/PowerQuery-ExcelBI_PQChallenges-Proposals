/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7030022607387176960-Bpqj
PQ proposal in 4 main steps
1. Add 1 dummy row with values : Min&Max of From-To, Amount=0, Item=First 1 (no matter)
    No to get 1 missing empty month, in final result. 
    =For the tricky requirement of Feb-22, ... 😉 
2. Transform data in each row to 1 new list of Records 
   for each day of period(From-To) with MonthName and dailyAmount 
3. Expand NewData in List 
4. Pivot by Month with Sum of DailyAmounts by Item 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData051"]}[Content],
    Rename = Table.RenameColumns(Source,{{"From Date", "From"}, {"To Date", "To"}}),
    ChangeType = Table.TransformColumnTypes(Rename, {{"From", type date}, {"To", type date}, {"Item", type text}, {"Amount", Int64.Type}}),
    AddRow_AllDates = Table.InsertRows(ChangeType, 0,
        {[From=List.Min(ChangeType[From]), To=List.Max(ChangeType[To]), Item=List.First(ChangeType[Item]), Amount=0]} ),
    Transform = Table.AddColumn(AddRow_AllDates, "Data", each let
       _ValPerDay = _[Amount]/(Duration.Days([To]-[From])+1),
       _RecordList = List.Transform({Number.From([From])..Number.From([To])}, 
          each [Month=Date.ToText(Date.From(_),[Format="MMM-yy",Culture="en-US"]), Val=_ValPerDay])
       in _RecordList),
    Expand = Table.ExpandRecordColumn(Table.ExpandListColumn(Transform[[Item], [Data]], "Data"), "Data", {"Month", "Val"}),
    Pivot = Table.Pivot(Expand, List.Distinct(Expand[Month]), "Month", "Val", each Number.Round(List.Sum(_&{0}),1))
in
    Pivot
