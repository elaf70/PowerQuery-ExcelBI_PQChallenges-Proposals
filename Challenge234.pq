let
  Source = Excel.CurrentWorkbook(){[Name="tData234"]}[Content],
  Transform = Table.FromRecords(Table.TransformRows(Source, each let 
    _Leaves = List.Split(List.Skip(List.RemoveNulls(Record.ToList(_))),2),
    _LeaveDays = List.Transform (_Leaves, each let
      _StartD = Date.From(_{0}), 
      _dates = List.Dates(_StartD, Duration.Days(Date.From(_{1}) - _StartD) + 1, #duration(1,0,0,0))
      in List.Select(_dates, each Date.DayOfWeek(_, Day.Monday) < 5) )
    in _[[Employee]] & [TotalLeaves=List.Count(List.Distinct(List.Combine(_LeaveDays)))] ) ),
  Sort = Table.Sort(Transform, "Employee")
in
  Sort
