/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7156858362087849984-I_yF

PQ proposal in 2 main steps
    1) Tranform
    2) Select Columns

For fun try to use as much as possible LIST functions only. 
Without other Table functions : Join/Expand/Combine/Group
==================================*/
let
    T2 = Table.Buffer(Excel.CurrentWorkbook(){[Name="tData151_2"]}[Content]),
    T1 = Table.TransformColumnTypes(Excel.CurrentWorkbook(){[Name="tData151_1"]}[Content],
        {{"Start Date", type date}, {"End Date", type date}}),
    Transform = Table.AddColumn(T1, "Total Hours", each  let
        _V = Record.FieldValues(_),
        _D = List.Dates([Start Date], Duration.Days([End Date]-[Start Date])+1, #duration(1,0,0,0)),
        _DD = List.TransformMany(_D, (x)=>{Date.DayOfWeek(x)}, (x,y)=>T2{y-1}),
        _T = Table.RemoveRowsWithErrors(Table.FromColumns({_D, _DD}, {"Date", "DD"})),
        Add_Hours = Table.AddColumn(_T, "Hours", each let
            _TS = List.Max({[DD][Start Time], if ([Date]=_V{1}) then _V{2} else null}),
            _TE = List.Min({[DD][End Time], if ([Date]=_V{3}) then _V{4} else null})
            in Duration.TotalHours(Time.From(_TE) - Time.From(_TS)) )
        in List.Sum(Add_Hours[Hours]) ),
    Result =Transform[[Employee], [Total Hours]] 
in
    Result



