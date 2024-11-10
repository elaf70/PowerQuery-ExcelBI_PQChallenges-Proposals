/*===============================
Ref = https://www.linkedin.com/feed/update/urn:li:activity:7261226065468018690
PQ proposal in 2 main-steps
1) Transform each row by splitting LeaveDates by 2 and or each find days then count the combination of all distinct ones 
2) Sort by [Employee]
=================================*/

let
  Source = Excel.CurrentWorkbook(){[Name="tData234"]}[Content],
  Transform = Table.FromRecords(Table.TransformRows(Source, each let 
    _Leaves = List.Split(List.Skip(List.RemoveNulls(Record.ToList(_))),2),
    _LeaveDays = List.Transform(_Leaves, each let
      _dates = List.DateTimes(_{0}, Number.From(_{1} - _{0}) + 1, #duration(1,0,0,0))
      in List.Select(_dates, each Date.DayOfWeek(_, Day.Monday) < 5) )
    in _[[Employee]] & [TotalLeaves=List.Count(List.Distinct(List.Combine(_LeaveDays)))] ) 
    ),
  Sort = Table.Sort(Transform, "Employee")
in
  Sort
