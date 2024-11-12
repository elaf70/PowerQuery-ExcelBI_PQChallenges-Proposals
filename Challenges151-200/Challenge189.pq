/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7205055956089544704-Nyjv

PQ proposal in 1 main steps
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData189"]}[Content],
    Result = Table.AddColumn(Table.SelectRows(Source, each [Code] <> "Yes"), "Result",
        each if (Source[Code]{[Seq]}?="Yes") then "Pass" else "Fail")
in
    Result


