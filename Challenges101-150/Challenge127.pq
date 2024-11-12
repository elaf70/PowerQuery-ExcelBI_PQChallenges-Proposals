/*================================
Ref = https://www.linkedin.com/posts/excelbi_powerquerychallenge-daxchallenge-m-activity-7126417783898185728-1ywM

PQ proposal in 2 main steps
    1) Set TableOfZodiacSign by Splitting + ChangeType for From/To Day items in Integer 
    2) Add ZodiacSign by Selecting Matching MonthName + correct rule for To/From Day 
==================================*/
let
    TZS = let
        _Source = Excel.CurrentWorkbook(){[Name="tData127_2"]}[Content],
        _Split = Table.SplitColumn(_Source, 
            "Zodiac Signs", 
            Splitter.SplitTextByEachDelimiter({" ("," ","– "," ",")"}), {"Sign", "FromM", "FromD", "ToM", "ToD"} )
        in Table.TransformColumnTypes(_Split,{{"FromD", Int64.Type}, {"ToD", Int64.Type}}),
        
    Source = Excel.CurrentWorkbook(){[Name="tData127"]}[Content],
    Add_ZS = Table.AddColumn(Source, "Zodiac Sign", each let
        _M = Date.MonthName([Date of Birth], "en-Us"), 
        _D = Date.Day([Date of Birth]),
        _S = Table.SelectRows(TZS, each (([FromM]=_M and _D>=[FromD]) or ([FromM]=_M and _D<=[ToD])) )
        in _S[Sign]{0})
in
    Add_ZS



