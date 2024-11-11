/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7055752147514105856-LuJt
PQ proposal in 4 main steps
    1) Split 
    2) Group by Dept : Accumulating all Data + CountRow 
    3) Create the list of NewDataColumnNames 
    4) Split AccumulatedData with its NewColumnsNames 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData072"]}[Content],
    Split_ANS = Table.SplitColumn(Source, "Age & Nationality & Salary", Splitter.SplitTextByDelimiter(",", QuoteStyle.None), {"Age", "Nationality", "Salary"}),
    Group = Table.Group(Split_ANS, {"Dept"}, {
      {"Data", each let
        _List = List.Accumulate(Table.ToRows(Table.RemoveColumns(_,"Dept")), {}, (s,a)=>s&a),
        _Combine = Text.Combine(List.Transform(_List, Text.From), ";")
        in _Combine},
      {"RowCount", each Table.RowCount(_)}
    }),
    NewColNames = List.Combine(
      List.Transform({1..List.Max(Group[RowCount])}, 
      (a)=> List.Transform({"Employee","Age","Nationality","Salary"}, (b)=> b & Text.From(a) ))),
    SplitData = Table.SplitColumn(Table.RemoveColumns(Group,"RowCount"), "Data",
      Splitter.SplitTextByDelimiter(";", QuoteStyle.Csv), NewColNames)
in
    SplitData


