/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7169541932719718400-rc9f

PQ proposal in 2 main steps
    1)Group +AddColumn with List funtions 
    2) Expand 
==================================*/
let
    fxYWAddWeeks = (yw, n)=>let 
        y = Number.IntegerDivide(yw,100), w=Number.Mod(yw,100),
        d = Date.AddWeeks(#date(y,1,1), w+n)
        in Date.Year(d)*100 + Date.WeekOfYear(d),
 
    Source = Excel.CurrentWorkbook(){[Name="tData161"]}[Content],
    Group = Table.Group(Source, "Group", {"All", each let
            _tb = Table.Buffer(_), 
            _WMin = fxYWAddWeeks (List.Min([Week No]), 6), 
            AddOcc = Table.AddColumn(_, "Max Occured No", each let
                _W = [Week No], _lW = {fxYWAddWeeks(_W, -8).._W}, 
                _t = Table.SelectRows(_tb, 
                    each (_W>=_WMin and List.Contains(_lW, [Week No]))),
                _l1 = List.Modes(_t[Winning No]),
                _l2 = List.Transform(List.Sort(List.Distinct(_l1)), Number.ToText)  
                in Text.Combine(_l2, ", ") )
            in AddOcc}),
    Expand = Table.ExpandTableColumn(Group, "All", {"Week No", "Winning No", "Max Occured No"})
in 
    Expand


