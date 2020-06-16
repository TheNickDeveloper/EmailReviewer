using EmailReviewer.BusinessLogic;
using EmailReviewer.Services;
using System.Configuration;
using Serilog;

namespace EmailReviewer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Establish logger
            Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(ConfigurationManager.AppSettings["LogFolderPath"].ToString()
                , rollingInterval: RollingInterval.Day
                , rollOnFileSizeLimit: true)
            .CreateLogger();

            Log.Information("Start email reviewing.");

            // Establish sqliteDb and table
            var conn = ConfigurationManager.ConnectionStrings["IncidentTicketContext"].ConnectionString;
            var sqliteHelper = new SqliteHelper(conn);
            sqliteHelper.CreateTable("IncidentTickets"
                , new string[] { "TicketId", "Priority" }
                , new string[] { "string", "string" });

            try
            {
                var incidentChecker = new IncidentChecker();
                incidentChecker.ReviewIncidentFromEmailFolder();
                Log.Information("Finish scripting.");
            }
            catch (System.Exception e)
            {
                Log.Error(e.Message);
            }

        }
    }
}
