/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7123881070973833217-uYJf

PQ proposal in 4 main steps
    1) ...
    2) ...
    3) ... 
    4) ...
==================================*/
let
    fxListEqual = (l1, l2)=> List.Accumulate(l1, [r=true, i=0],
       (s,c)=>[r=s[r] and (c=l2{s[i]} or c="*"), i=s[i]+1])[r],

    TKO = Excel.CurrentWorkbook(){[Name="tData125_2"]}[Content],
    LKO = List.Transform(Table.ToRows(TKO), 
        each {List.RemoveLastN(_), List.Last(_)}) ),
    Source = Excel.CurrentWorkbook(){[Name="tData125"]}[Content],
    Add_O = Table.AddColumn(Source, "Output", each
        List.Accumulate(LKO, null, (s,c)=> 
            if(s=null and fxListEqual(c{0},Record.ToList(_))) then c{1} else s))
in
   Add_O


