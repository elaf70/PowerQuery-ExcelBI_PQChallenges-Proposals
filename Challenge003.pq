/*===============================
Ref = https://www.linkedin.com/posts/excelbi_powerbi-powerquery-dax-activity-6969964727720398849-5Bc-
PQ proposal in 3 main-steps
1) Sort all distinct letter from column [Alphabet] 
2) Transform this list to extract for each all its occurences in initial Source
3) Create result table in columns with it 
=================================*/

let
  Source = Excel.CurrentWorkbook(){[Name="tData003"]}[Content],
  BufferValues=List.Buffer(Source[Alphabets]),
  DistinctLetters = List.Sort(List.Distinct(BufferValues)),
  ListsByLetters = List.Transform(DistinctLetters, each let 
    Letter = _,
    in List.Select(BufferValues, each _=Letter) ),
  Result = Table.FromColumns(ListsByLetters, DistinctLetters)
in
  Result
