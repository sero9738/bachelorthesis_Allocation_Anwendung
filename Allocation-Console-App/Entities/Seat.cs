namespace Allocation_Console_App.Entities
{
    public class Seat
    {
        public int SeatNumber { get; set; }
        public Table? Parent { get; set; }
        public Group? Owner { get; set; }
        public bool Occupied { get; set; }
        public int Score { get; set; }
        public int col { get; set; } = -1;
        public int row { get; set; } = -1;
    }
}