/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7136927031631429632-_PE5

PQ proposal in 3 main steps
    1) Group {[ID]} per [Group] with also TableBuffer for further calls
    2) Transform rows to requested Selection to avoid previous Group and only keep those with List.Intersect on IDs 
    3) ToTable from rows 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData137"]}[Content],
    Group = Table.Buffer(Table.Group(Source, {"Group"}, {"IDs", each [ID]})),
    Transform = List.Transform(Table.ToRecords(Group), (r)=> let
        _LCG = Table.SelectRows(Group, each ([Group]>r[Group]) and (List.IsEmpty(List.Intersect({[IDs],r[IDs]}))=false))[Group]
        in {r[Group],Text.Combine(_LCG, ",")} ),
    ToTable = Table.FromRows(Transform, {"Group", "Common"})
in
    ToTable


