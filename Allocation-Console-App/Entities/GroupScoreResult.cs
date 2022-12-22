namespace Allocation_Console_App.Entities
{
    public class GroupScoreResult
    {
        public Group Parent { get; set; }
        public int AverageScore { get; set; }
        public List<Seat> OwnedSeats { get; set; }
        public Seat MainSeat { get; set; }

        public GroupScoreResult(Group parent, List<Seat> ownedSeats)
        {
            Parent = parent;
            OwnedSeats = ownedSeats;
            AverageScore = GetAverageSeatScore(ownedSeats, parent);
            MainSeat = GetMainSeatFromList(ownedSeats);
        }

        public int GetAverageSeatScore(List<Seat> seats, Group group)
        {
            int sumValue = 0;
            foreach (var seat in seats)
            {
                if (seat.Owner != null && seat.Owner == group)
                {
                    sumValue += seat.Score;
                }
            }
            return (sumValue / seats.Count);
        }

        public Seat GetMainSeatFromList(List<Seat> seats)
        {
            Seat mainSeat = null;
            foreach (var seat in seats)
            {
                if (mainSeat == null || mainSeat.Score < seat.Score)
                {
                    mainSeat = seat;
                }
            }
            return mainSeat;
        }
    }
}
