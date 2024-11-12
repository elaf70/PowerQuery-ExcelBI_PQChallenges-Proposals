/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7129355402814574592-vLFd

PQ proposal in 2 main steps
    1) Transform T2 by adding a new [No] column from unpivoting 2 columns + add col of all N° of T1 items included
    2) Select expected mapping columns + sort 
==================================*/
let
    tVal = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData130"]}[Content]),
    tEmp = Excel.CurrentWorkbook(){[Name="tData130_2"]}[Content],
    transform = let 
        unpivot = Table.Unpivot(tEmp, {"City / Country","Text"}, "Column Name","Value"),
        add_No = Table.AddColumn(unpivot, "No", each let
            v = [Value]
            in Table.SelectRows(tVal, each Text.PositionOf(v, [Value], 0, Comparer.OrdinalIgnoreCase)>=0)[No] ),
        expand = Table.ExpandListColumn(add_No, "No")
        in Table.SelectRows(expand, each [No] <> null),
    selectColAndSort = Table.Sort(transform[[Emp ID],[No],[Column Name]], {"Emp ID", "No"})
in
    selectColAndSort


