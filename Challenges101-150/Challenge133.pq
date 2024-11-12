/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7134032291835617280-x2-A
PQ proposal in 3 main steps
    1) Merge
    2) Transfom rows
    3) Add 1 row for GranTotal
==================================*/

//VERSION 2 : 1 peu + "optimisée" = avec 1 seule passe & 1 seule lecture de chaque lignes  
//            (No Pivot/Unpivot/Sort/Group ....)   NB : moins lisible & comprehensible directement

let
    DurTSec = Duration.TotalSeconds,
    fxGetValues=(pVal,cVal,coef)=>List.Transform(List.Zip({pVal,cVal}), each Number.Round((_{1}-_{0})*coef)),

    Source = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name="tData133"]}[Content],
        {{"Time", type time}, {"Date", type date}}),
    MCN = List.Skip(Table.ColumnNames(Source),2),
    Merge = Table.CombineColumns(Source, MCN, each _, "Values"),
    Transform = let
        T = List.Accumulate(Table.ToRows(Merge), [r={}, cum={0,0,0}, prev=null], (s,c)=>let
            p = s[prev]??c
            in if (c{0}=p{0}) 
                then [r=s[r], cum=List.Transform(List.Zip({s[cum], fxGetValues(p{2},c{2},1)}), List.Sum), prev=c]
                else let 
                    coef = DurTSec(#time(24,0,0)-p{1}) / DurTSec((#time(24,0,0)-p{1})+(c{1}-#time(0,0,0))),
                    c1 = List.Transform(List.Zip({s[cum], fxGetValues(p{2},c{2},coef)}), List.Sum),
                    c2 = fxGetValues(p{2},c{2},1-coef)
                    in [r=s[r] & {{p{0}} & c1 & {List.Sum(c1)}}, cum=c2, prev=c]),
        M = T[r] & {{T[prev]{0}} & T[cum] & {List.Sum(T[cum])}}
        in Table.FromRows(M, {"Date"} & MCN & {"Total Consumption"}), 
    Add_RGT = let 
        RGT = {"Grand Total"} & List.Transform(List.Skip(Table.ToColumns(Transform)), List.Sum)
        in Table.FromRows(Table.ToRows(Transform) & {RGT}, Table.ColumnNames(Transform)) 
in 
    Add_RGT


