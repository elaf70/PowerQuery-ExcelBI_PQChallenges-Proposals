/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7220638624931848193-YT_y

PQ proposal in 2 main steps
    1) Transform rows using List.Accumulate
        With Serial calculated from previous & col. position of name 
    2) Get final result table from these transformed rows
==================================*/
let
    fxNextSerial = (s)=>let
        xHeadTrail = [p=Text.PositionOf(s,".", Occurrence.Last), 
            Head =if p=-1 then "" else Text.Start(s,p), 
            Trail=if p=-1 then s else Text.Middle(s,p+1)]
        in xHeadTrail[Head] & (if xHeadTrail[Head]<>"" then "." else "") & Text.From(Number.From(xHeadTrail[Trail])+1),
    Source = Excel.CurrentWorkbook(){[Name="tData202"]}[Content],
    Transform= List.Accumulate(Table.ToRows(Source), [i=0, x="0", r={}], (s,c)=>let
        i = List.PositionOf(c, List.RemoveNulls(c){0}),
        x = if (i>s[i]) then s[x]&".1" 
            else if i=s[i] then fxNextSerial(s[x])
            else fxNextSerial(Text.Start(s[x], Text.PositionOf(s[x], ".", Occurrence.All){i}))
        in [i=i, x=x, r=s[r]&{{x,c{i}}}] ),
    Result = Table.FromRows(Transform[r], {"Serial", "Names"})
in
    Result

