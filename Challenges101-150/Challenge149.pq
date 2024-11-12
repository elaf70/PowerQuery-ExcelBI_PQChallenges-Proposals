/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7154322067520118784-XQ-0

PQ proposal in 3 main steps
    1) Set DateType for Start and End dates
    2) Add new column to split into table with 1 single List.Generate
        each raw period [Start]=>[End] by all EndOfMonth(s) in between 
        with new [Per Diem] value, based on NbOfDays of each split 
    3) Expand inner-tables 
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData149"]}[Content],
    CType = Table.TransformColumnTypes(Source,{{"Start Date", type date}, {"End Date", type date}}),
    Add_Split = Table.AddColumn(CType, "Split", each let 
        _S = _[Start Date], _E = _[End Date], _P=_[Per Diem],
        L = List.Generate(
            ()=>[S=Date.StartOfMonth(_S),E=Date.AddDays(_S,-1),P=0], 
            each _[S]<_E, 
            each let 
                newS = Date.AddDays(_[E],+1), 
                newE = List.Min({Date.EndOfMonth(newS), _E})
                in [S=newS, E=newE, P=(Duration.Days(newE-newS)+1)*_P] )
        in Table.FromRecords(List.Skip(L))),
    Expand = Table.ExpandTableColumn(Add_Split[[Employee],[Split]], "Split", 
        {"S", "E", "P"}, List.Skip(Table.ColumnNames(Source)) )
in
    Expand


