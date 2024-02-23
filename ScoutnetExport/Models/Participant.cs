namespace ScoutnetExport.Models
{
    public sealed class Participant
    {
        public int Id { get; set; }

        public string Groupname { get; set; }

        public string PatrolName { get; set; }

        public string MemberNumber { get; set; }

        public string Name { get; set; }

        public string Gender { get; set; }

        public int ZipCode { get; set; }

        public DateOnly BirhtDate { get; set; }

        public string Role { get; set; }

        public Dictionary<DateOnly, bool> ParticipationDates { get; set; } = new Dictionary<DateOnly, bool>();
    }
}
