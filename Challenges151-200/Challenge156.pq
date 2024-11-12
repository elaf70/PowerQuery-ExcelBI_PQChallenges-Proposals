/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7162294178624724992-_l5y

PQ proposal in 3 main steps with Record usage + Group 

As for FUN 😀 , I was looking for sth else than my 1st thought & more usual steps in such case with : TableJoin / Pivot / Row+Col reorders 
 ==================================*/
let
    Source = Excel.CurrentWorkbook(), 
    Teachers = Record.FromTable(
        Table.RenameColumns(Source{[Name="tData156_2"]}[Content],
        {{"Subjects","Name"},{"Teacher","Value"}}) ), 
    Group = Table.Group(Source{[Name="tData156"]}[Content], {"Name"}, 
        {"All", each Record.SelectFields(Teachers, [Subjects])} ),
    Expand = Table.ExpandRecordColumn(Group, "All", Record.FieldNames(Teachers))
in
    Expand



