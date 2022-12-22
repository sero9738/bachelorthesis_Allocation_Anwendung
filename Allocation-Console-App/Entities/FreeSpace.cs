namespace Allocation_Console_App.Entities
{
    public class FreeSpace
    {
        public Table? Parent { get; set; }
        public int StartingIndex { get; set; }
        public int Size { get; set; }
        public Seat BestSeat { get; set; }
        public int Score { get; set; }

        public int SetScore()
        {
            int score = 0;
            if (Parent != null)
            {
                for (int i = StartingIndex; i < (StartingIndex + Size); i++)
                {
                    score += Parent.Seats[i].Score;
                }
                score = score / Size;
            }
            return score;
        }

        public bool CheckIfAllSeatsAreFree()
        {
            if (Parent != null)
            {
                for (int i = StartingIndex; i < (StartingIndex + Size); i++)
                {
                    if (Parent.Seats[i].Occupied)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public List<Seat> GetSeats(int groupSize)
        {
            List<Seat> seats = new List<Seat>();
            if (groupSize <= Size && Parent != null && CheckIfAllSeatsAreFree())
            {
                if (groupSize == Size)
                {
                    for (int i = StartingIndex; i < (StartingIndex + Size); i++)
                    {
                        seats.Add(Parent.Seats[i]);
                    }
                }
                else
                {
                    int bottomCount;
                    int topCount;
                    int mainSeatIndex = Parent.Seats.IndexOf(BestSeat);
                    if (mainSeatIndex >= StartingIndex && mainSeatIndex < (StartingIndex + Size))
                    {
                        bottomCount = mainSeatIndex - StartingIndex;
                        topCount = (StartingIndex + Size) - mainSeatIndex;
                        int range = groupSize / 2;
                        if (groupSize == 1)
                        {
                            if (topCount == 0)
                            {
                                seats.Add(Parent.Seats[StartingIndex + Size]);
                            }
                            else if (bottomCount == 0)
                            {
                                seats.Add(Parent.Seats[StartingIndex]);
                            }
                            else
                            {
                                seats.Add(Parent.Seats[mainSeatIndex]);
                            }
                        }
                        else if (range > topCount)
                        {
                            for (int i = ((StartingIndex + Size) - groupSize); i < (StartingIndex + Size); i++)
                            {
                                seats.Add(Parent.Seats[i]);
                            }
                        }
                        else if (range > bottomCount)
                        {
                            for (int i = StartingIndex; i < (StartingIndex + groupSize); i++)
                            {
                                seats.Add(Parent.Seats[i]);
                            }
                        }
                        else
                        {
                            for (int i = ((mainSeatIndex + range) - groupSize); i < (mainSeatIndex + range); i++)
                            {
                                seats.Add(Parent.Seats[i]);
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine($"GetSeats failed for group.");
            }
            return seats;
        }

    }
}
