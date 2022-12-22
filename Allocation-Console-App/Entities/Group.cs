namespace Allocation_Console_App.Entities
{
    public class Group
    {
        public int GroupId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Day { get; set; } = string.Empty;
        public int Size { get; set; }
        public int Score { get; set; }
        public bool Placed { get; set; } = false;
        public int ColorIndex { get; set; } = 0;

    }
}