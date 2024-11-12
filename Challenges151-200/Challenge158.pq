/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7164830890236358656-wCXZ

PQ proposal in 2 main steps
    (1) Zip 2 fist header rows
    (2) Transform each other rows on List of records splitting on New Dept
    (3) Convert Records to table
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData158"]}[Content],
    Rows = Table.ToRows(Source), 
    CN = {"Group", "Dept"} & List.Skip(List.Distinct(Rows{1})),
    H = List.Zip(List.FirstN(Rows,2)),
    Records = List.Accumulate(List.RemoveFirstN(Rows,2), {}, (s,r)=>let
        Transform = List.Accumulate(List.Skip(List.Zip({H,r})), 
            [lr={}, cr=[Dept="", #"Emp ID"=null]], (s,c)=>let
                _New = if (c{0}{0}=null) then [lr={}, cr=s[cr]] else [lr={s[cr]}, cr=[Group=r{0}, Dept=c{0}{0}]] 
                in [lr=s[lr] & _New[lr], cr=Record.AddField(_New[cr], c{0}{1}, c{1})] )
        in s & Transform[lr] & {Transform[cr]} ),
    ToTable = Table.FromRecords(List.Select(Records, each [Emp ID]<>null), CN, MissingField.UseNull)
in
    ToTable


