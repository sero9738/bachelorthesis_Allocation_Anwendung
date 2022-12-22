using Allocation_Console_App.Entities;

namespace Allocation_Console_App
{
    public class AllocationService
    {
        public AllocationContext Start(AllocationContext context, bool isFriday)
        {
            // Teilaufgabe 1 - Sortierung
            context.OrderByScore();
            List<GroupScoreResult> results = new List<GroupScoreResult>();
            List<Group> orderedGroupData;
            if (isFriday)
            {
                orderedGroupData = context.OrderedGroupDataFriday;
            }
            else
            {
                orderedGroupData = context.OrderedGroupDataSaturday;
            }
            foreach (var group in orderedGroupData)
            {
                // Teilaufgabe 5 - Freie Bereiche ermitteln.             
                // Liste von Bereichen nach besten Sitzplatz sortieren. So kann über jeden Bereich iteriert werden bis ein Bereich gefunden wurde in dem die aktuelle Gruppe platz hat.
                List<FreeSpace> freeSpaceList = new();
                if (isFriday)
                {
                    foreach (var table in context.TheaterFriday)
                    {
                        freeSpaceList.AddRange(table.IdentifieFreeSpaces());
                    }
                }
                else
                {
                    foreach (var table in context.TheaterSaturday)
                    {
                        freeSpaceList.AddRange(table.IdentifieFreeSpaces());
                    }
                }
                // Teilaufgabe 4 - Finden den besten verfügbaren Sitzplatz.
                freeSpaceList = freeSpaceList.OrderByDescending(x => x.Score).ToList();

                // Iteriere über alle freien Bereiche bist du einen freien Bereich gefunden hast in den die Gruppe reinpasst.
                for (int i = 0; i < freeSpaceList.Count && !group.Placed; i++)
                {
                    if (group.Size < freeSpaceList[i].Size)
                    {
                        // Teilaufgabe 6 - Platzieren um den Hauptsitz; Platziere Gruppe in um den besten Platz im Feld
                        var result = freeSpaceList[i].Parent.PlaceGroupAroundMainSeat(group, freeSpaceList[i]);
                        if (result != null)
                        {
                            results.Add(result);
                        }
                        group.Placed = true;
                    }
                }
            }
            if (isFriday)
            {
                context.AllocationResultsFriday = results;
            }
            else
            {
                context.AllocationResultsSaturday = results;
            }
            return context;
        }

        // Teilaufgabe 9 - Fasse Gruppen zu Blöcken zusammen um eine gerade Größe zu erhalten
        private List<GroupBlock> BlockGroups(List<Group>? groups)
        {
            List<GroupBlock> blocks = new();
            GroupBlock block;
            if (groups == null || !groups.Any())
            {
                return blocks;
            }
            groups = groups.OrderByDescending(x => x.Score).ToList();
            bool match = true;
            while (match)
            {
                block = new();
                match = false;
                for (int i = 0; i < groups.Count && !match; i++)
                {
                    // Durchsuche Liste nach Gruppen mit ungeraden größen.
                    if (groups[i].Size % 2 == 1)
                    {
                        // Wenn gefunden, erzeuge neuen GroupBlock. Füge hinzu zur Liste.
                        block.Groups.Add(groups[i]);
                        block.TotalSize += groups[i].Size;
                        // Wenn Block eine gerade Gesamtgröße hat, speichere diesen in einer Liste und beginne die Suche von vorne
                        if (block.TotalSize % 2 == 0)
                        {
                            blocks.Add(block);
                            match = true;
                        }
                    }
                }
            }
            return blocks;
        }

        public int GetMaxScore(List<GroupScoreResult> results)
        {
            int maxScore = 0;
            foreach (var result in results)
            {
                // Mutipliziere den Score mit der Gewichtung der Gruppe. Um so höher das Ergebnis um so optimaler das Ergebnis.
                maxScore += result.AverageScore * result.Parent.Score;
            }
            return maxScore;
        }
    }
}
