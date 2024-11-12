/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7177152088689307648-dOfs

PQ proposal in 3 main steps
    1) Get a Sorted list of all names
    2) Group by campain to create for each a table with all name's results
    3) Combine inner tables
==================================*/
let
  Source = Excel.CurrentWorkbook(){[Name="tData167"]}[Content],
  N = List.Sort(List.Distinct(Table.SelectRows(Source[[Camp No],[Vaccine]], each ([Camp No] is number)=false)[Vaccine])),
  Group = Table.Group(Source, "Camp No",
    {"G", each let
        V = [Vaccine]{0}, 
        C = Table.ToColumns(Table.RemoveFirstN(_,1)),
        fxValue = (col,p, alt)=>try if (col{p} is datetime) then "Yes" else "No" otherwise alt
        in List.Transform(N, each let p=List.PositionOf(N,_) in {p+1, V, _, fxValue(C{2},p,"No"),fxValue(C{3},p,"N/A")}) 
    }, 0, (p,c)=>Number.From(c is number)),
  Expand = Table.FromRows(List.Combine(Group[G]), {"Camp No", "Vaccine", "Name", "Notified", "Administred"})
in
  Expand



