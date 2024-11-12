/*================================
Ref = https://www.linkedin.com/posts/excelbi_challenge-powerquerychallenge-daxchallenge-activity-7192372368030232577-zv0e

PQ proposal in 3 main steps
    1) Group by [Team] to create a table of Players + List of HighScorersNames + HighestScore
    2) Find PlayersColumnNames from the one with MaxPlayers
    3) Expand TableColumn of Players
==================================*/
let
    Source = Excel.CurrentWorkbook(){[Name="tData179"]}[Content],
    HNames = {"Highest Scoring Player", "Highest Score"},
    Group = Table.Group(Source, "Team", {
            {"G", each let
                pList = [Player],
                SBest = Table.Sort(_, {"Runs Scored", Order.Descending}){0},
                CName = {"Team"} & List.Transform({1..List.Count(pList)}, each "Player"&Text.From(_)) & HNames
                in Table.FromRows({{SBest[Team]} & pList & {SBest[Player], SBest[Runs Scored]}}, CName)
            },
            { "NBP", each Table.RowCount(_)}
            }),
    PosMaxP = List.PositionOf(Group[NBP], List.Max(Group[NBP])),
    Combine = Table.Combine(Group[G], Table.ColumnNames(Group[G]{PosMaxP}))
in
    Combine


