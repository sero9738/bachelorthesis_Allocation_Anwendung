namespace Allocation_Console_App.Entities
{
    public class Table
    {
        public int TableNumber { get; set; }
        public List<Seat> Seats { get; set; } = new List<Seat>();

        public int CountFreeSeats()
        {
            int count = 0;
            foreach (var seat in Seats)
            {
                if (!seat.Occupied)
                {
                    count++;
                }
            }
            return count;
        }

        // Teilaufgabe 5 - Freie Bereiche identifizieren
        public List<FreeSpace> IdentifieFreeSpaces()
        {
            List<FreeSpace> freeSpaces = new List<FreeSpace>();
            int spaceStartingIndex = -1;
            int spaceSize = 0;
            foreach (Seat seat in Seats)
            {
                if (!seat.Occupied)
                {
                    if (spaceStartingIndex < 0)
                    {
                        spaceStartingIndex = Seats.IndexOf(seat);
                    }
                    spaceSize++;
                }
                else
                {
                    if (spaceStartingIndex >= 0)
                    {
                        FreeSpace space = new();
                        space.Parent = this;
                        space.StartingIndex = spaceStartingIndex;
                        space.Size = spaceSize;

                        // Teilaufgabe 4 - Identifzieren des besten Sitzplatzes
                        for (int i = space.StartingIndex; i < space.StartingIndex + space.Size; i++)
                        {
                            if (space.BestSeat == null || (space.BestSeat.Score < Seats[i].Score && !Seats[i].Occupied))
                            {
                                space.BestSeat = Seats[i];
                            }
                        }
                        space.Score = space.SetScore();
                        freeSpaces.Add(space);
                        spaceStartingIndex = -1;
                        spaceSize = 0;
                    }
                }
            }

            if (spaceSize > 0)
            {
                FreeSpace space = new();
                space.Parent = this;
                space.StartingIndex = spaceStartingIndex;
                space.Size = spaceSize;
                // Speichere besten Platz innerhalb des Bereiches.
                for (int i = space.StartingIndex; i < space.StartingIndex + space.Size; i++)
                {
                    if (space.BestSeat == null || (space.BestSeat.Score < Seats[i].Score && !Seats[i].Occupied))
                    {
                        space.BestSeat = Seats[i];
                    }
                }
                space.Score = space.SetScore();
                freeSpaces.Add(space);
            }

            return freeSpaces;
        }

        // Teilaufgabe 7 - Verschieben von Gruppen (Richtung Bühne)
        public void MoveFieldsToTop(int seatNumber, int spacesToMove)
        {
            var fields = IdentifieFreeSpaces();
            int mainIndex = Seats.FindIndex(x => x.SeatNumber == seatNumber);
            if (mainIndex >= 0 && mainIndex < Seats.Count && Seats[mainIndex].Occupied)
            {
                FreeSpace spaceAbove = null;
                foreach (var space in fields)
                {
                    if (space.StartingIndex + space.Size < mainIndex)
                    {
                        spaceAbove = space;
                    }
                }
                if (spaceAbove != null && spacesToMove < spaceAbove.Size)
                {
                    int endIndex = spaceAbove.StartingIndex + spaceAbove.Size - 1;
                    while (spacesToMove > 0)
                    {
                        int currentIndex = endIndex;
                        while (Seats[currentIndex + 1].Occupied && Seats[currentIndex + 1].Owner != null)
                        {
                            Seats[currentIndex].Occupied = true;
                            Seats[currentIndex].Owner = Seats[currentIndex - 1].Owner;

                            Seats[currentIndex + 1].Occupied = false;
                            Seats[currentIndex + 1].Owner = null;

                            currentIndex++;
                        }
                        spacesToMove--;
                        endIndex--;
                    }
                }
            }
        }


        // Teilaufgabe 7 - Verschieben von Gruppen (Richtung Türen)
        public void MoveFieldsToBottom(int seatNumber, int spacesToMove)
        {
            var fields = IdentifieFreeSpaces();
            int mainIndex = Seats.FindIndex(x => x.SeatNumber == seatNumber);
            if (mainIndex >= 0 && mainIndex < Seats.Count && Seats[mainIndex].Occupied)
            {
                FreeSpace spaceBelow = null;
                foreach (var space in fields)
                {
                    if (spaceBelow == null && space.StartingIndex > mainIndex || spaceBelow != null && spaceBelow.StartingIndex <= space.StartingIndex)
                    {
                        spaceBelow = space;
                    }
                }
                if (spaceBelow != null && spacesToMove < spaceBelow.Size)
                {
                    int endIndex = spaceBelow.StartingIndex;
                    while (spacesToMove > 0)
                    {
                        int currentIndex = endIndex;
                        while (Seats[currentIndex - 1].Occupied && Seats[currentIndex - 1].Owner != null)
                        {
                            Seats[currentIndex].Occupied = true;
                            Seats[currentIndex].Owner = Seats[currentIndex - 1].Owner;

                            Seats[currentIndex - 1].Occupied = false;
                            Seats[currentIndex - 1].Owner = null;

                            currentIndex--;
                        }
                        spacesToMove--;
                        endIndex++;
                    }
                }
            }
        }

        // Teilaufgabe 6 - Platzieren einer Gruppe um einen Hauptsitz
        public GroupScoreResult? PlaceGroupAroundMainSeat(Group group, FreeSpace space)
        {
            if (group != null && space != null)
            {
                var seats = space.GetSeats(group.Size);
                if (seats != null && seats.Any())
                {
                    List<Seat> owned = new();

                    foreach (var seat in seats)
                    {
                        seat.Occupied = true;
                        seat.Owner = group;
                        owned.Add(seat);
                    }

                    GroupScoreResult result = new(group, owned);
                    return result;
                }
            }
            return null;
        }

        // Teilaufgabe 8 - Analyse und Evaluierung von Lösungen
        public List<GroupScoreResult> EvaluteTable()
        {
            List<GroupScoreResult> scores = new();

            Group? currentGroup = null;
            List<Seat> ownedSeats = new();

            foreach (var seat in Seats)
            {
                if (currentGroup != null && seat.Owner != null && seat.Owner.GroupId == currentGroup.GroupId)
                {
                    ownedSeats.Add(seat);
                }
                else
                {
                    if (ownedSeats.Any() && currentGroup != null)
                    {
                        GroupScoreResult result = new(currentGroup, ownedSeats);
                        scores.Add(result);
                        ownedSeats = new();
                        currentGroup = null;
                    }
                    if (currentGroup == null && seat.Owner != null)
                    {
                        currentGroup = seat.Owner;
                        ownedSeats.Add(seat);
                    }
                }
            }
            return scores;
        }
    }
}