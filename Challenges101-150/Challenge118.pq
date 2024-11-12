/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7114096596434194432-adoX

PQ proposal in 2 main steps  using List functions
    1) Transform with List.Accumulate to create new records
    2) Creat Result table with it
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData118"]}[Content],
    Transform = List.Accumulate( Table.ToRecords(Source),       [Done = {}, Index = 0],       (s,c)=> let
         _d2 = List.RemoveNulls({c[Dept1], c[Dept2]}),
         _Select = List.Select(s[Done],              each [Mail]=c[Mail] or List.Contains(_d2, [Dept1]) or List.Contains(_d2, [Dept2])),
         _i = if List.IsEmpty(_Select) then s[Index]+1 else _Select{0}[Index]
         in [ Done = s[Done]&{Record.AddField(c, "Index", _i)}, Index = List.Max({_i, s[Index]})]     ),
    ToTable = Table.FromRecords(Transform[Done])
in
    ToTable


