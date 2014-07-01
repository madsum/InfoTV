using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.DirectoryServices.AccountManagement;
using System.Diagnostics;              
using System.Globalization;
using System.Runtime.InteropServices;
using System.Web.Security;                                                                                                                                                                                                                                                                                                                                                                                                                                                            
using System.Reflection;
using System.Security.Principal;
using System.Web;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Xml;
using System.DirectoryServices;

namespace OutlookCalenderFetcher
{
    public class OutlookCalender
    {
        public List<Meeting> TodayMeetingList {get; set;}
        public List<Meeting> TomorrowMeetingList{get; set;}
        private string Username = null;
        private string Passwored = null;
        private string AutodiscoverEmail = null;
        public string ErrorMsg = null;
        public  int RefreshSec { get; set; }

        public OutlookCalender()
        {

        }

        public void Initialize(string username, string password, string autodiscoverEmail)
        {
            Username = username;
            Passwored = password;
            AutodiscoverEmail = autodiscoverEmail;
            TodayMeetingList = new List<Meeting>();
            TomorrowMeetingList = new List<Meeting>();
            //ContactDic = new Dictionary<string, string>();
            //FilteredContact = new Dictionary<string, string>();
        }
    
        public bool GetAllData()
        {
            //Dictionary<string, string>.KeyCollection keys = GlobalProperty.MeetinRooms.Keys;

            try
            {
                foreach (KeyValuePair<string, string> MeetingKeyVal in GlobalProperty.MeetinRooms)
                {
                    if (!GetAllMeets(MeetingKeyVal.Value, MeetingTime.Today))
                    {
                        return false;
                    }
                    if (!GetAllMeets(MeetingKeyVal.Value, MeetingTime.Tomorrow))
                    {
                        return false;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }

            if (SortList())
            {
                return true;
            }
            else
            {
                return false;
            }
        }
                   

        private bool SortList()
        {
            try
            {
                TodayMeetingList = TodayMeetingList.OrderBy(x => x.KloDate).ToList();
                TomorrowMeetingList = TomorrowMeetingList.OrderBy(x => x.KloDate).ToList();
                return true;
            }
            catch (System.Exception e)
            {
                return false;
            }

        }

   

        // this enum used for get meeting time today and tomorrow
        enum MeetingTime
        {
            Today,
            Tomorrow
        };
  
        // It find a keywrod in the input_str it return short predefined string 
        private string ContinasCaseInsensitiv(string input_str)
        {
            CultureInfo culture = new CultureInfo("sv-FI");
            string output_str = null;

            if (String.IsNullOrEmpty(input_str))
            {
                return output_str;
            }
            if (culture.CompareInfo.IndexOf(input_str, ShortName.Akku, CompareOptions.IgnoreCase) >= 0)
            {
                output_str = ShortName.Akku;
                return output_str;
            }
            else if (culture.CompareInfo.IndexOf(input_str, ShortName.Iso, CompareOptions.IgnoreCase) >= 0)
            {
                output_str = ShortName.Iso;
                return output_str;
            }
            else if (culture.CompareInfo.IndexOf(input_str, ShortName.Laturi, CompareOptions.IgnoreCase) >= 0)
            {
                output_str = ShortName.Laturi;
                return output_str;
            }
            else if (culture.CompareInfo.IndexOf(input_str, ShortName.Pikku, CompareOptions.IgnoreCase) >= 0)
            {
                output_str = ShortName.Pikku;
                return output_str;
            }
            else if (culture.CompareInfo.IndexOf(input_str, ShortName.Pikku2, CompareOptions.IgnoreCase) >= 0)
            {
                output_str = ShortName.Pikku2;
                return output_str;
            }

            output_str = input_str;
            return output_str;
        }

        private void AppendMeeting(FindItemsResults<Appointment> appointments, string email, MeetingTime mTime )
        {
            if (mTime == MeetingTime.Today)
            {
                foreach (Appointment a in appointments)
                {
                    Meeting apt = new Meeting();
                    apt.Aihe = a.Subject;
                    apt.Isanta = a.Organizer.ToString(); ;
                    apt.Paivamaara = a.Start.ToShortDateString();
                    apt.KloDate = a.Start;
                    apt.Klo = a.Start.ToString("HH:mm");
                    // normally location is a huge string so kepp it shorter
                    apt.Paikka = ContinasCaseInsensitiv(a.Location);
                    TodayMeetingList.Add(apt);
                }
            }
            else if (mTime == MeetingTime.Tomorrow)
            {
                foreach (Appointment a in appointments)
                {
                    Meeting apt = new Meeting();
                    apt.Aihe = a.Subject;
                    apt.Isanta = a.Organizer.ToString(); ;
                    apt.Paivamaara = a.Start.ToShortDateString();
                    apt.KloDate = a.Start;
                    apt.Klo = a.Start.ToString("HH:mm");
                    // normally location is a huge string so kepp it shorter
                    apt.Paikka = ContinasCaseInsensitiv(a.Location);
                    TomorrowMeetingList.Add(apt);
                }
            }
        }
                
        private bool GetAllMeets(string email, MeetingTime mTime)
         {
             ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
             service.Credentials = new WebCredentials(Username, Passwored, GlobalProperty.Domain);
             try
             {
                service.AutodiscoverUrl(AutodiscoverEmail, RedirectionUrlValidationCallback);
                Mailbox principle = new Mailbox(email);

                if (mTime == MeetingTime.Today)
                {
                    DateTime startDate = DateTime.Today;
                    DateTime endDate = DateTime.Today.AddDays(1);
                    Microsoft.Exchange.WebServices.Data.CalendarView cView = new Microsoft.Exchange.WebServices.Data.CalendarView(startDate, endDate);

                    try
                    {
                        CalendarFolder calendar = CalendarFolder.Bind(service, new FolderId(WellKnownFolderName.Calendar, principle), new PropertySet());
                        // Retrieve a collection of appointments by using the calendar view.
                        FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);
                        AppendMeeting(appointments, email, mTime);
                    }
                    catch
                    {
                        Debug.WriteLine("ignore this");
                    }
                }
                else if (mTime == MeetingTime.Tomorrow)
                {
                    DateTime startDate = DateTime.Today.AddDays(1);
                    DateTime endDate = DateTime.Today.AddDays(2);
                    Microsoft.Exchange.WebServices.Data.CalendarView cView = new Microsoft.Exchange.WebServices.Data.CalendarView(startDate, endDate);
                    try
                    {
                        CalendarFolder calendar = CalendarFolder.Bind(service, new FolderId(WellKnownFolderName.Calendar, principle), new PropertySet());
                        // Retrieve a collection of appointments by using the calendar view.
                        FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);
                        AppendMeeting(appointments, email, mTime);
                    }
                    catch
                    {
                        Debug.WriteLine("ignore this");
                    }
                }
             }
             catch (System.Exception )
             {
                 return false;
             }
             return true;
         }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        // This are room know know by all.
        static class ShortName
        {
            // all Mikkelä room
            public static string Laturi = "Laturi";
            public static string Pikku = " Pikku paristo";
            public static string Pikku2 = " Pikkuparisto";
            public static string Iso = "Iso-Paristo";
            public static string Akku = "Akku";
            // all Mankka room
            public static string LogPieni = "Mankkaa LOG Pieni";
            public static string LogIso = "Mankkaa LOG Iso";
        }
    }

    // this calls represent a meeting
    public class Meeting
    {
        public string Aihe { get; set; }
        public string Isanta { get; set; }
        public string Paivamaara { get; set; }
        public string Klo { get; set; }
        public DateTime KloDate { get; set; }
        public string Paikka { get; set; }
        public string Kutsutut { get; set; }
        public string Lisatiedot { get; set; }
    }
}
