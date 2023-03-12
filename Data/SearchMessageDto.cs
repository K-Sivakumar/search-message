namespace SearchMessage.Data
{
    public class SearchMessageDto
    {
        public DateTime ReceivedDate { get; set; }

        public string? ReceivedFrom { get; set; }

        public string Summary { get; set; }

        public string? Subject { get; set; }

        public int? Rank { get; set; }
    }
}
