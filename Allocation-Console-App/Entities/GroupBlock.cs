namespace Allocation_Console_App.Entities
{
    public class GroupBlock
    {
        public List<Group> Groups { get; set; } = new();
        public int TotalSize { get; set; } = 0;
    }
}