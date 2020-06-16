using System.ComponentModel.DataAnnotations;

namespace EmailReviewer.DbContexts.Models
{
    public class IncidentTicket
    {
        [Key]
        public string TicketId { get; set; }
        public string Priority { get; set; }
    }
}
