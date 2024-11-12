/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7213027975884111872-fiOU

PQ proposal in 3 main steps
    1) Create Sorted-List of [subjects] + complete ColumNames to use at end 
    2) Table.Group by [Class] to create for each subset-row a record value 
    3) Table.FromRecords to get expected result 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData196"]}[Content],
    Subjects = List.Sort(List.Distinct(Source[Subject])), 
    CN = Subjects & List.Transform(Subjects, each "Marks-" & _),
    Group = Table.Group(Source, "Class", 
        {"G", each let
            _T = Table.TransformRows(_, each 
                Record.FromList({[Class], [Marks]},{[Subject], "Marks-" & [Subject]}))
            in Record.Combine(_T)
        } ),
    Result = Table.FromRecords(Group[G], CN, MissingField.UseNull)
in
    Result

