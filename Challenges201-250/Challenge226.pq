/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7251079210427944961-DRbG

PQ proposal in 4 main steps
    1) FillDown [DepID] for missing values 
    2) Group by [Dept] to create sorted tables of all related Employees
    3) Expand these inner-Tables 
    4) ChangeType : mainly for proper PromotionDate 
==================================*/
let
    CN = {"Emp Names", "Promotion Date", "Salary"}, 
    Source = Excel.CurrentWorkbook(){[Name="tData226"]}[Content],
    FillDown = Table.FillDown(Source,{"Dept ID"}),
    Group = Table.Group(FillDown, {"Dept ID"}, {"G", each let
        _Values = List.RemoveNulls([Promoted Employees] & [Not Promoted Employees]), 
        _T = Table.FromColumns({_Values},{"ESP"}),
        _Split = Table.SplitColumn(_T, "ESP", Splitter.SplitTextByDelimiter("-"), {CN{0},CN{2},CN{1}}),
        _Sort = Table.Sort(_Split, CN{0})
        in _Sort }),
    Expand = Table.ExpandTableColumn(Group, "G", CN),
    ChgType = Table.TransformColumnTypes(Expand,{{CN{1}, type date}, {CN{2}, Int64.Type} })
in
    ChgType

/*
// VERSION 2 : UNPIVOT
//====================
let

    Source = Excel.CurrentWorkbook(){[Name="tData226"]}[Content],
    PH = Table.ColumnNames(Source),
    NH = {"Dept ID", "Emp Names", "Promotion Date", "Salary"}, 
    FillDown = Table.FillDown(Source, {PH{0}}),
    UnPivot = Table.Unpivot(FillDown, {PH{2}, PH{3}}, "Attribute", "Value"),
    Split = Table.SplitColumn(UnPivot, "Value", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {NH{1},NH{3},NH{2}}),
    SelectCols = Table.SelectColumns(Split,NH),
    ChgType = Table.TransformColumnTypes(SelectCols,{{NH{2}, type date},{NH{3}, Int64.Type}}),
    Sort = Table.Sort(ChgType, {NH{0}, NH{1}})
in
    Sort

*/
