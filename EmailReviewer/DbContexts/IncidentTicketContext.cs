using EmailReviewer.DbContexts.Models;
using System.Data.Entity;

namespace EmailReviewer.DbContexts
{
    public class IncidentTicketContext : DbContext
    {
        public DbSet<IncidentTicket> IncidentTickets { get; set; }
    }
}
