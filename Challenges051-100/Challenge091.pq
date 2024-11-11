/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7080756915927982080-xnDV
PQ proposal in 4 main steps
    1) Transpose Store+Date items 
    2) Transpose Customer
    3) Add Repeat items (longuest part) 
    4) Reorder columns 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData091"]}[Content],
    Tr_SV = List.Accumulate({{"Store", Table.FillDown, type text}, {"Visit Date", Table.FillUp, type date}}, 
    Source, (s,c)=>let
        _AddCol = Table.AddColumn(s, c{0}, each if ([Data1]=c{0}) then [Data2] else null, c{2}),
        _CType = Table.TransformColumnTypes(_AddCol,{c{0}, c{2}})
        in Function.Invoke(c{1}, {_CType, {c{0}}}) ),
    Tr_C = let
        _T = Table.SelectRows(Tr_SV, each [Data1]="Customers"),
        _Group = Table.Group(_T, {"Store", "Visit Date"},
            {"Customer", each Text.Split(Text.Combine(_[Data2], ", "), ", ") })
        in Table.ExpandListColumn(_Group, "Customer"),
    Add_Repeat = let
        _G = Table.Group(Tr_C, {"Customer"}, {"D", each _}),
        _LRV = List.Accumulate(_G[D], {}, 
            (s,c)=>if Table.RowCount(c)=1 then s else let
                _VI = Table.AddIndexColumn(Table.Sort(c, "Visit Date"), "VI", 1),
                _R = Table.AddColumn(_VI, "Repeat", (r) => let
                _s = Table.SelectRows(_VI, each [VI] = r[VI] - 1)[Store]{0}?
                in if (_s=null) then "R" else if (_s=r[Store]) then "RS" else "RD")
            in s & Table.ToRecords(_R) ),
        _CN = {"Customer","Visit Date"},
        _J = Table.AddJoinColumn(Tr_C, _CN, Table.FromRecords(_LRV), _CN, "Repeat")
        in Table.ExpandTableColumn(_J, "Repeat", {"Repeat"}),
  ReOrder = Table.SelectColumns(Add_Repeat, {"Store", "Customer", "Visit Date", "Repeat"})
in
  ReOrder


