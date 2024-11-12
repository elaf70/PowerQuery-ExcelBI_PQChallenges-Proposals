/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7187298942521393153-gI9N

PQ proposal in 3 main steps
    1) Add column with NextGen N° (+1)
    2) Groupby [Family] and make an inner join with NextGen N° + combine GenN° columns Into Relationship + Reorder columns
    3) Combine sub-tables & Sort
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData175"]}[Content],
    AddNextG = Table.AddColumn(Source,"NextG", each [#"Generation No"]+1),
    GroupF = Table.Group(AddNextG, {"Family"}, {{"G", each let
        Merge = Table.NestedJoin(_, "NextG", _, "Generation No", "Merge", JoinKind.Inner),
        ExpandNextG = Table.ExpandTableColumn(Merge, "Merge", {"Name"}, {"Next Generation"}),
        CombineGNum = Table.CombineColumns(ExpandNextG, {"Generation No", "NextG"},
            (x)=>Text.From(x{0}) & " - " & Text.From(x{1}), "Relationship")
        in CombineGNum[[Name],[Family],[Next Generation],[Relationship]]
    }})[G],
    CombineSort = Table.Sort(Table.Combine(GroupF), {"Family", "Relationship", "Name", "Next Generation"})
in
    CombineSort

/*
SOLUTION2 (Zoran Milocanovic)
let
    Source = Excel.CurrentWorkbook(){[Name="tData175"]}[Content],
    RowsBuff = List.Buffer(Table.ToRows(Source)), 
    H = {"Name", "Family", "Next Generation", "Relationship"},
    Transform = List.TransformMany(RowsBuff, 
        each List.Select(RowsBuff, (p) => p{2} = _{2} + 1 and p{1} = _{1}), 
        (i, _) => {i{0}, i{1}} & {_{0}, Text.From(i{2}) & " - " & Text.From(_{2})}),
    ToTable = Table.FromRows(Transform, H),    
    Sort = Table.Sort(ToTable, {H{1}, H{3}, H{0}, H{2}})
in
 Sort
*/

