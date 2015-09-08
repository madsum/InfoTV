using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalenderFetcher
{
    public static class GlobalProperty
    {
        public const string EncryptionPassword = "mypass12";
        public const string LDAP = "LDAP://espoo.ad.city/DC=espoo,DC=ad,DC=city";
        public const string Domain = "Espoo";
        public const string AutodiscoverUrl = "https://mail.espoo.fi/EWS/Exchange.asmx";

        public static string Username { get; set; }
        public static string Password { get; set; }
        //public static string Email { get; set; }
        //public static string AutodiscoverEmail { get; set; }
        //public static string LDAP { get; set; }
        //public static string Domain { get; set; }
        //public static string EncryptionPassword { get; set; }
        public static bool Authenticated { get; set; }
        public static string Message { get; set; }      

        public static OutlookCalender Calender = new OutlookCalender();
        public static Dictionary<string, string> MeetinRooms = new Dictionary<string, string>(); 
    }
}
