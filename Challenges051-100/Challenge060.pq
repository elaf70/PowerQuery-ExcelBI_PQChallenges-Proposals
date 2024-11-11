/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7040531855405658112-neZD
PQ proposal in 4 main steps
    1. Transform Data into Records from 2 TextSplit + apply the 2 rules
        ItemName: if any remove final 's' / ItemType: if Nothing=>"New" 
    2. Build 1 table from the RecordsList 
    3. Pivot on ItemType(New/Used) 
    4. Add Total Column 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData060"]}[Content],
    Transform = Table.TransformColumns(Source, {{"Data", each let
        _Records = List.Transform(Text.Split(_, "/"), each let 
             _l = Text.Split(Text.Trim(_), " "),
             _Item = if(Text.EndsWith(_l{1},"s")) then Text.RemoveRange(_l{1}, Text.Length(_l{1})-1) else _l{1},
             _Type= if(List.Count(_l)=3) then _l{2} else "New"
             in [Items=_Item, Type=_Type, Count=Number.From(_l{0})] )
        in _Records }}),
    Expand = Table.FromRecords(Table.ExpandListColumn(Transform, "Data")[Data], 
        type table[Items=Text.Type,Type=Text.Type,Count=Number.Type]),
    Pivot = Table.Pivot(Expand, List.Distinct(Expand[Type]), "Type", "Count", List.Sum),
    Add_Total = Table.AddColumn(Pivot, "Total", each List.Sum({[New],[Used]}))
in
    Add_Total


