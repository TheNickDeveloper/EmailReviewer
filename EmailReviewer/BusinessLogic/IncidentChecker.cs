using EmailReviewer.DbContexts;
using EmailReviewer.DbContexts.Models;
using EmailReviewer.Exceptions;
using Microsoft.Office.Interop.Outlook;
using Serilog;
using System.Linq;

namespace EmailReviewer.BusinessLogic
{
    public class IncidentChecker
    {
        private readonly Application _outlookApp;
        private readonly IncidentTicketContext _incidentTicketContext;
        private readonly MAPIFolder _inboxFolder;

        private MAPIFolder _itNotificationFolder;
        private MAPIFolder _incidentsCriticalFolder;
        private MAPIFolder _incidentsNormalFolder;
        private MAPIFolder _othersFolder;

        public MAPIFolder ItNotificationFolder
        {
            get
            {
                if (_itNotificationFolder != null)
                {
                    return _itNotificationFolder;
                }
                else
                {
                    try
                    {
                        _itNotificationFolder = _inboxFolder.Folders["ItNotification"];
                    }
                    catch (System.Exception)
                    {
                        throw new EmailFolderNoFoundException("Cannot find ItNotification folder.");
                    }
                    return _itNotificationFolder;
                }
            }
        }

        public MAPIFolder IncidentsCriticalFolder
        {
            get
            {
                if (_incidentsCriticalFolder != null)
                {
                    return _incidentsCriticalFolder;
                }
                else
                {
                    foreach (MAPIFolder folder in _itNotificationFolder.Folders)
                    {
                        if (folder.Name == "Incidents_Critical")
                        {
                            _incidentsCriticalFolder = folder;
                            break;
                        }
                    }

                    if (_incidentsCriticalFolder == null)
                    {
                        _itNotificationFolder.Folders.Add("Incidents_Critical");
                        _incidentsCriticalFolder = _itNotificationFolder.Folders["Incidents_Critical"];
                    }

                    return _incidentsCriticalFolder;
                }
            }
        }

        public MAPIFolder IncidentsNormalFolder
        {
            get
            {
                if (_incidentsNormalFolder != null)
                {
                    return _incidentsNormalFolder;
                }
                else
                {
                    foreach (MAPIFolder folder in _itNotificationFolder.Folders)
                    {
                        if (folder.Name == "Incidents_Normal")
                        {
                            _incidentsNormalFolder = folder;
                            break;
                        }
                    }

                    if (_incidentsNormalFolder == null)
                    {
                        _itNotificationFolder.Folders.Add("Incidents_Normal");
                        _incidentsNormalFolder = _itNotificationFolder.Folders["Incidents_Normal"];
                    }

                    return _itNotificationFolder;
                }
            }
        }

        public MAPIFolder OthersFolder
        {
            get
            {
                if (_othersFolder != null)
                {
                    return _othersFolder;
                }
                else
                {
                    foreach (MAPIFolder folder in _itNotificationFolder.Folders)
                    {
                        if (folder.Name == "Others")
                        {
                            _othersFolder = folder;
                            break;
                        }
                    }

                    if (_othersFolder == null)
                    {
                        _itNotificationFolder.Folders.Add("Others");
                        _othersFolder = _itNotificationFolder.Folders["Others"];
                    }

                    return _othersFolder;
                }
            }
        }

        public IncidentChecker()
        {
            _outlookApp = new Application();
            _incidentTicketContext = new IncidentTicketContext();
            _inboxFolder = _outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        }

        public void ReviewIncidentFromEmailFolder()
        {
            Items inboxMails = ItNotificationFolder.Items;

            do
            {
                if (inboxMails.Count == 0)
                {
                    Log.Information("No new mail in ItNotification folder.");
                    return;
                }

                // rearrange mail item ascending by date
                inboxMails = ItNotificationFolder.Items;
                inboxMails.Sort("[Subject]", true);

                var currentMail = inboxMails[1] as MailItem;
                var subjectCategory = currentMail.Subject.Substring(0, 8);
                if (subjectCategory == "Incident")
                {
                    var priority = currentMail.Subject.Substring(9, 2);
                    var incidentCateogty = GetIncidentCategory(priority);
                    string incidentNumber;

                    switch (incidentCateogty)
                    {
                        case "IncidentWithPriority":
                            MoveToFolder(priority, currentMail);
                            AutoReply(priority, currentMail);
                            incidentNumber = currentMail.Subject.Substring(12, 10);
                            RecordIncidentNumberIntoSqliteDb(incidentNumber, priority);
                            break;

                        case "CommandOrIncidentWithoutPriority":
                            incidentNumber = currentMail.Subject.Substring(9, 10);
                            priority = GetPriorityFromDb(incidentNumber);
                            MoveToFolder(priority, currentMail);
                            break;
                    }
                }
                else
                {
                    Log.Information($"Not incident notice, move mail:{currentMail.Subject} to Otherd folder.");
                    currentMail.Move(OthersFolder);
                }

            } while (inboxMails.Count != 0);
        }

        private string GetIncidentCategory(string priority)
        {
            return (priority == "P1" || priority == "P2" || priority == "P3" || priority == "P4" || priority == "P5")
                ? "IncidentWithPriority"
                : "CommandOrIncidentWithoutPriority";
        }

        private void MoveToFolder(string priority, MailItem currentMail)
        {
            if (priority == "P4")
            {
                currentMail.Move(IncidentsCriticalFolder);
                var folderName = "IncidentCritial Folder";
                Log.Information($"Move mail:{currentMail.Subject} to {folderName}.");
            }

            if (priority == "P5")
            {
                currentMail.Move(IncidentsNormalFolder);
                var folderName = "IncidentNormal Folder";
                Log.Information($"Move mail:{currentMail.Subject} to {folderName}.");
            }

            if (priority == "NoDefined")
            {
                currentMail.Move(OthersFolder);
                var folderName = "Others Folder";
                Log.Information($"Cannot recognize priority, move mail:{currentMail.Subject} to {folderName}.");
            }
        }

        private void AutoReply(string priority, MailItem currentMail)
        {
            if (priority == "P4")
            {
                var newMail = currentMail.Forward();
                newMail.To = "nick.tsai@effem.com";
                string contents = "hi, this is forward test <br/> <br/>";
                newMail.HTMLBody = contents + newMail.HTMLBody;
                newMail.Save();
                Log.Information($"Classify as Critial Mail, execute auto reply to {newMail.To}.");
            }
        }

        private void RecordIncidentNumberIntoSqliteDb(string incidentNumber, string priority)
        {
            var newIncidentTicket = new IncidentTicket { TicketId = incidentNumber, Priority = priority };
            var allTickets = _incidentTicketContext.IncidentTickets.ToList();
            var exsitingTickets = allTickets.Find(x => x.TicketId == newIncidentTicket.TicketId);

            if (exsitingTickets == null)
            {
                _incidentTicketContext.IncidentTickets.Add(newIncidentTicket);
                _incidentTicketContext.SaveChanges();
            }

            Log.Information($"Store incident number:{incidentNumber} to DB.");
        }

        private string GetPriorityFromDb(string ticketId)
        {
            var tickets = _incidentTicketContext.IncidentTickets.ToList();
            var target = tickets.Find(x => x.TicketId == ticketId);

            if (target != null)
            {
                return target.Priority;
            }

            return "NoDefined";
        }

    }
}
