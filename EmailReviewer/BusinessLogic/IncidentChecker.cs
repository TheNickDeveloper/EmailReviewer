using EmailReviewer.DbContexts;
using EmailReviewer.DbContexts.Models;
using EmailReviewer.Exceptions;
using Microsoft.Office.Interop.Outlook;
using Serilog;
using System.Configuration;
using System.Linq;

namespace EmailReviewer.BusinessLogic
{
    public class IncidentChecker
    {
        private readonly Application _outlookApp;
        private readonly IncidentTicketContext _incidentTicketContext;
        private readonly MAPIFolder _inboxFolder;
        private MAPIFolder _targetSourceFolder;
        private MAPIFolder _incidentsCriticalFolder;
        private MAPIFolder _incidentsNormalFolder;
        private MAPIFolder _othersFolder;
        private MAPIFolder _meetingFodler;
        private MAPIFolder _appointmentFolder;

        public IncidentChecker()
        {
            _outlookApp = new Application();
            _inboxFolder = _outlookApp.GetNamespace("MAPI").GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            _incidentTicketContext = new IncidentTicketContext();
        }

        public MAPIFolder SourceTargetFolder
        {
            get
            {
                var folderName = ConfigurationManager.AppSettings["TargetSourceFolder"].ToString();

                if (_targetSourceFolder != null)
                {
                    return _targetSourceFolder;
                }
                else
                {
                    try
                    {
                        if (folderName == "Inbox")
                        {
                            _targetSourceFolder = _inboxFolder;
                        }
                        else
                        {
                            _targetSourceFolder = _inboxFolder.Folders[folderName];
                        }
                    }
                    catch (System.Exception)
                    {
                        throw new EmailFolderNoFoundException($"Cannot find {folderName} folder.");
                    }
                    return _targetSourceFolder;
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
                    foreach (MAPIFolder folder in _targetSourceFolder.Folders)
                    {
                        if (folder.Name == "Incidents_Critical")
                        {
                            _incidentsCriticalFolder = folder;
                            break;
                        }
                    }

                    if (_incidentsCriticalFolder == null)
                    {
                        _targetSourceFolder.Folders.Add("Incidents_Critical");
                        _incidentsCriticalFolder = _targetSourceFolder.Folders["Incidents_Critical"];
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
                    foreach (MAPIFolder folder in _targetSourceFolder.Folders)
                    {
                        if (folder.Name == "Incidents_Normal")
                        {
                            _incidentsNormalFolder = folder;
                            break;
                        }
                    }

                    if (_incidentsNormalFolder == null)
                    {
                        _targetSourceFolder.Folders.Add("Incidents_Normal");
                        _incidentsNormalFolder = _targetSourceFolder.Folders["Incidents_Normal"];
                    }

                    return _incidentsNormalFolder;
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
                    foreach (MAPIFolder folder in _targetSourceFolder.Folders)
                    {
                        if (folder.Name == "Others")
                        {
                            _othersFolder = folder;
                            break;
                        }
                    }

                    if (_othersFolder == null)
                    {
                        _targetSourceFolder.Folders.Add("Others");
                        _othersFolder = _targetSourceFolder.Folders["Others"];
                    }

                    return _othersFolder;
                }
            }
        }

        public MAPIFolder MeetingFolder
        {
            get
            {
                if (_meetingFodler != null)
                {
                    return _meetingFodler;
                }
                else
                {
                    foreach (MAPIFolder folder in _targetSourceFolder.Folders)
                    {
                        if (folder.Name == "Meeting")
                        {
                            _meetingFodler = folder;
                            break;
                        }
                    }

                    if (_meetingFodler == null)
                    {
                        _targetSourceFolder.Folders.Add("Meeting");
                        _meetingFodler = _targetSourceFolder.Folders["Meeting"];
                    }

                    return _meetingFodler;
                }
            }
        }

        public MAPIFolder AppointmentFolder
        {
            get
            {
                if (_appointmentFolder != null)
                {
                    return _appointmentFolder;
                }
                else
                {
                    foreach (MAPIFolder folder in _targetSourceFolder.Folders)
                    {
                        if (folder.Name == "Appointment")
                        {
                            _appointmentFolder = folder;
                            break;
                        }
                    }

                    if (_appointmentFolder == null)
                    {
                        _targetSourceFolder.Folders.Add("Appointment");
                        _appointmentFolder = _targetSourceFolder.Folders["Appointment"];
                    }

                    return _appointmentFolder;
                }
            }
        }

        public void ReviewIncidentFromEmailFolder()
        {
            Items inboxMails = SourceTargetFolder.Items;
            
            do
            {
                if (inboxMails.Count == 0)
                {
                    Log.Information("No new mail in ItNotification folder.");
                    return;
                }

                // Have to assigin the value again, ow it wont update automatically.
                inboxMails = SourceTargetFolder.Items;
                inboxMails.Sort("[Subject]", true);

                // mail item
                if (inboxMails[1] is MailItem currentMail) MailItemAction(currentMail);

                // meeting item
                if (inboxMails[1] is MeetingItem meetingMail) MeetingItemAction(meetingMail);

                // appointment item
                if (inboxMails[1] is AppointmentItem appointmentMail) AppointmentItemAction(appointmentMail);

            } while (inboxMails.Count != 0);
        }

        private void MailItemAction(MailItem currentMail)
        {
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
        }

        private void MeetingItemAction(MeetingItem meetingMail)
        {
            meetingMail.Move(MeetingFolder);
            var folderName = "Meeting Folder";
            Log.Information($"Move meeting:{meetingMail.Subject} to {folderName}.");
        }

        private void AppointmentItemAction(AppointmentItem appointmentMail)
        {
            appointmentMail.Move(AppointmentFolder);
            var folderName = "Appointment Folder";
            Log.Information($"Move appointment:{appointmentMail.Subject} to {folderName}.");
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
                newMail.Subject = "Testing";
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
            return target != null ? target.Priority : "NoDefined";
        }

    }
}
