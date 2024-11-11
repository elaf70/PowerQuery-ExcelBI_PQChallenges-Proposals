/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-6994869799080112129-ZkXb
PQ proposal in 3 main steps
1) ...
2) ...
3) ... 
==================================*/
let
    //--- Get 1 List of Names from Custumers + Keys values for Table t  (Source / By Group of OrderN°)
    fxGetLCK = (t)=>let
      _LC= List.Sort(List.Distinct(t[Customer])),
      _LKVal = List.Distinct(Table.SelectRows(t, each ([Key] = "X" or [Key] = "Y"))[Customer]),
      _LK = List.Sort(List.Transform(_LKVal,each _&" Key"))   
      in _LC & _LK,

    Source = Excel.CurrentWorkbook(){[Name="tData024"]}[Content],
    Group = Table.Group(Source, {"Order Nr"}, 
       {"Data", each List.Accumulate(fxGetLCK(_), [], (s,c)=>Record.AddField(s,c,1))} ),
    Expand = Table.ExpandRecordColumn(Group, "Data", fxGetLCK(Source)),
    NullToZero = Table.TransformColumns(Expand,{},each if _=null then 0 else _)
in
    NullToZero

