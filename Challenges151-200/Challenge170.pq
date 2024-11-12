/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7180051191526002688-0CgR

PQ proposal in 4 main steps
    1) Add DayOfWeek, 
    2) Group By DayOfWeek and then by Item to create Record of Lowest/Highest Sales in a Record 
    3) Expand record
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData170"]}[Content],
    Add_DoW = Table.AddColumn(Table.TransformColumnTypes(Source,{{"Date", type date}}),"DoW", 
        each if (Date.DayOfWeek([Date],Day.Monday)>=5) then "Weekend" else "Weekday"),
    Group = Table.Group(Add_DoW, "DoW", {"G", each let 
        G = Table.Sort(Table.Group(_, "Item", {"Sales",each List.Sum([Sale])}), "Sales"),
        _fx = (lf)=>Text.Combine(List.Reverse(Table.SelectRows(G, each [Sales]=lf(G[Sales]))[Item]), ", ")
        in [TS=List.Sum([Sale]), HSI=_fx(List.Last), LSI=_fx(List.First)] }),
    Expand = Table.ExpandRecordColumn(Group, "G", {"TS", "HSI", "LSI"}, {"Total Sales", "Highest Selling Item", "Lowest Selling Item"})
in
    Expand


