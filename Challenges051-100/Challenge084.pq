/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7070972435927306241-H-iX
PQ proposal in 3 main steps
    (1) Transform With List.Accumulate 
    (2) FilterRows 
    (3)Group
 ==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData084"]}[Content],
    Transform = List.Accumulate({"Type", "Sub type"}, Source, 
        (s,c)=> let
            _AddCol = Table.AddColumn(s, c, each  if [Classification]=c then [Value] else null)
            in Table.FillDown(_AddCol, {c}) ),
    FilterRows = Table.SelectRows(Transform, each ([Classification] <> "Type")),
    Group = Table.Group(FilterRows, {"Company", "Type", "Sub type"}, 
        {"Price", each try Table.SelectRows(_, each ([Classification]="Price"))[Value]{0} otherwise null} )
in
    Group


