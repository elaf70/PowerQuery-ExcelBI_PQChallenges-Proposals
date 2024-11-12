/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7101050629451624449-dUZM

PQ proposal in 3 main steps
    1. Combine values into 1 single GlobalList removing Null
    2. Split & Transform = Group/Calculate each ResultRow values 
    3. Convert to table (+optional ChgDateType=for easier display in XL)
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData107"]}[Content],
    ToList = List.RemoveNulls(List.Combine(Table.ToRows(Source))),
    SplitAndTransform = List.Transform(List.Split(ToList, 11), each {_{3}, _{4}, _{6}, _{8}+_{10}}),
    ToTable = Table.FromRows(SplitAndTransform, {"Date", "PO Number", "Supplier", "PO Value"}),
    ChangeType = Table.TransformColumnTypes(ToTable,{{"Date", type date}})
in
    ChangeType



