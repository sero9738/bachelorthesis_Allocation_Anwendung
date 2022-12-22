namespace Allocation_Console_App.Entities
{
    public class AllocationContext
    {
        public string SourceDirectory { get; set; } = string.Empty;
        public List<Group>? RawGroupData { get; set; }
        public List<Group>? OrderedGroupDataFriday { get; set; }
        public List<Group>? OrderedGroupDataSaturday { get; set; }
        public List<Table>? TheaterFriday { get; set; }
        public List<Table>? TheaterSaturday { get; set; }
        public List<GroupScoreResult>? AllocationResultsFriday { get; set; }
        public List<GroupScoreResult>? AllocationResultsSaturday { get; set; }
        public List<GroupScoreResult>? EvaluationResultsFriday { get; set; }
        public List<GroupScoreResult>? EvaluationResultsSaturday { get; set; }

        public void ProcessRawGroupData()
        {
            if (RawGroupData != null)
            {
                OrderedGroupDataFriday = new List<Group>();
                OrderedGroupDataSaturday = new List<Group>();
                foreach (var group in RawGroupData)
                {
                    if (group.Day.Equals("Freitag"))
                    {
                        OrderedGroupDataFriday.Add(group);
                    }
                    else if (group.Day.Equals("Samstag"))
                    {
                        OrderedGroupDataSaturday.Add(group);
                    }
                }
                OrderedGroupDataFriday = OrderedGroupDataFriday.OrderByDescending(group => group.GroupId).ToList();
                OrderedGroupDataSaturday = OrderedGroupDataSaturday.OrderByDescending(group => group.GroupId).ToList();
            }
            else
            {
                OrderedGroupDataFriday = null;
                OrderedGroupDataSaturday = null;
            }
        }

        public void OrderByScore()
        {
            if (OrderedGroupDataFriday != null && OrderedGroupDataSaturday != null)
            {
                OrderedGroupDataFriday = OrderedGroupDataFriday.OrderByDescending(group => group.Score).ToList();
                OrderedGroupDataSaturday = OrderedGroupDataSaturday.OrderByDescending(group => group.Score).ToList();
            }
        }

        public List<int> GetAllocatedSeats()
        {
            List<int> seatNumbers = new();
            foreach (var result in AllocationResultsFriday)
            {
                foreach (var seat in result.OwnedSeats)
                {
                    seatNumbers.Add(seat.SeatNumber);
                }
            }
            return seatNumbers;
        }
    }
}
