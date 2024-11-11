/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7048142001803792384-rNT6
PQ proposal in 2 main steps
    1) Calculate TotalsList GroupBy Product with 1 List.Generate
    2) Merge OriginalColumns values + TotalsValues
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData066"]}[Content],
    TotalByGroup = Table.Group(Source, {"Product"}, 
        {{"Total", each let 
            lr = Table.ToRecords(_),
            lt = List.Generate(
                ()=>[RTA=if (lr{0}[Category]="A") then lr{0}[Weight] else 0, RTB=0, T=0, i=0],
                each [i] < List.Count(lr),
                each let 
                    r = lr{[i]+1}
                    in [RTA = [RTA] + (if (r[Category]="A") then r[Weight] else 0), 
                        RTB = [RTB] + (if (RTA>0 and r[Category]="B") then r[Weight] else 0), 
                        T = if (r[Category]="B" and ((RTA - RTB) > 0)) then RTA - RTB else 0, 
                        i = [i] + 1 ],
                each [T])
            in lt }} ),
    Result = Table.FromColumns( 
        Table.ToColumns(Source) & {List.Combine(TotalByGroup[Total])}, 
        Table.ColumnNames(Source) & {"Total"} )
in
    Result


