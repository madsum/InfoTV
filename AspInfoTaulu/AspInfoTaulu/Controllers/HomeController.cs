using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using AspInfoTaulu.Models;
using OutlookCalenderFetcher;
using System.Web.Routing;
using System.Web.UI;
using System.IO;
using System.Security.Cryptography;
using System.Net;
using System.Configuration;
using HtmlAgilityPack;
using System.Web.Configuration;

namespace AspInfoTaulu.Controllers
{
    public class HomeController : Controller
    {
        private static Data mData = new Data();
        static string EncryptedFile = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/encrypted.txt");
        static Configuration RootWebConfig = WebConfigurationManager.OpenWebConfiguration("/appSettings");

        public ActionResult Index(string msg)
        {
            mData.Message = msg;

            if (!String.IsNullOrEmpty(msg))
            {
                return View(mData);
            }

            if (! ReadConfig())
            {
                return RedirectToAction("Tervetuloa", "Home");           
            }

            if (DecryptFile())
            {
                return RedirectToAction("Tervetuloa", "Home");           
            }
            return View(mData);
        }

        private bool ReadConfig()
        {
            GlobalProperty.MeetinRooms.Clear();

            if (RootWebConfig.AppSettings.Settings.Count > 0)
            {
                System.Configuration.KeyValueConfigurationElement EPass =
                RootWebConfig.AppSettings.Settings["EPass"];
                if (EPass != null)
                {
                    GlobalProperty.EncryptionPassword = EPass.Value;
                }

                System.Configuration.KeyValueConfigurationElement LDAP =
                RootWebConfig.AppSettings.Settings["LDAP"];
                if (LDAP != null)
                {
                    GlobalProperty.LDAP = LDAP.Value;
                }

                System.Configuration.KeyValueConfigurationElement Domain =
                RootWebConfig.AppSettings.Settings["Domain"];
                if (Domain != null)
                {
                    GlobalProperty.Domain = Domain.Value;
                }

                System.Configuration.KeyValueConfigurationElement Email =
                RootWebConfig.AppSettings.Settings["Email"];
                if (Domain != null)
                {
                    GlobalProperty.AutodiscoverEmail = Email.Value;
                }

                System.Configuration.KeyValueConfigurationElement Laturi =
                                    RootWebConfig.AppSettings.Settings["Laturi"];
                if (Laturi != null)
                {
                    GlobalProperty.MeetinRooms.Add(Laturi.Key, Laturi.Value);
                }

                System.Configuration.KeyValueConfigurationElement Pikku =
                RootWebConfig.AppSettings.Settings["Pikku paristo"];
                if (Pikku != null)
                {
                    GlobalProperty.MeetinRooms.Add(Pikku.Key, Pikku.Value);
                }

                System.Configuration.KeyValueConfigurationElement Iso =
                RootWebConfig.AppSettings.Settings["Iso paristo"];
                if (Iso != null)
                {
                    GlobalProperty.MeetinRooms.Add(Iso.Key, Iso.Value);
                }

                System.Configuration.KeyValueConfigurationElement Akku =
                RootWebConfig.AppSettings.Settings["Akku"];
                if (Akku != null)
                {
                    GlobalProperty.MeetinRooms.Add(Akku.Key, Akku.Value);
                }

                return true;
            }
            else
            {
                return false;
            }
        
        }

        private bool ReadDomainConfig3453()
        {
            //System.Configuration.Configuration rootWebConfig =
            //System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("/appSettings");

            if (RootWebConfig.AppSettings.Settings.Count > 0)
            {
                System.Configuration.KeyValueConfigurationElement EPass =
                RootWebConfig.AppSettings.Settings["EPass"];
                if (EPass != null)
                {
                    GlobalProperty.EncryptionPassword = EPass.Value;
                }

                System.Configuration.KeyValueConfigurationElement LDAP =
                RootWebConfig.AppSettings.Settings["LDAP"];
                if (LDAP != null)
                {
                    GlobalProperty.LDAP = LDAP.Value;
                }

                System.Configuration.KeyValueConfigurationElement Domain =
                RootWebConfig.AppSettings.Settings["Domain"];
                if (Domain != null)
                {
                    GlobalProperty.Domain = Domain.Value;
                }

                System.Configuration.KeyValueConfigurationElement Email =
                RootWebConfig.AppSettings.Settings["Email"];
                if (Domain != null)
                {
                    GlobalProperty.AutodiscoverEmail = Email.Value;
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool ReadRoomConfig435()
        { 
           // System.Configuration.Configuration rootWebConfig =
           // System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("/appSettings");

            if (RootWebConfig.AppSettings.Settings.Count > 0)
            {
                System.Configuration.KeyValueConfigurationElement Laturi =
                    RootWebConfig.AppSettings.Settings["Laturi"];
                if (Laturi != null)
                {
                    GlobalProperty.MeetinRooms.Add(Laturi.Key, Laturi.Value);
                }

                System.Configuration.KeyValueConfigurationElement Pikku =
                RootWebConfig.AppSettings.Settings["Pikku paristo"];
                if (Pikku != null)
                {
                    GlobalProperty.MeetinRooms.Add(Pikku.Key, Pikku.Value);
                }

                System.Configuration.KeyValueConfigurationElement Iso =
                RootWebConfig.AppSettings.Settings["Iso paristo"];
                if (Iso != null)
                {
                    GlobalProperty.MeetinRooms.Add(Iso.Key, Iso.Value);
                }

                System.Configuration.KeyValueConfigurationElement Akku =
                RootWebConfig.AppSettings.Settings["Akku"];
                if (Akku != null)
                {
                    GlobalProperty.MeetinRooms.Add(Akku.Key, Akku.Value);
                }
                return true;
            }
            else
            {
                return false;
            }
        }


        private bool EncryptFile(string inputdata)
        {
            try
            {
                using (FileStream fs = new FileStream(EncryptedFile, FileMode.Truncate)) 
                {
                    // empty any contact in the encrypted file
                }

                UnicodeEncoding UE = new UnicodeEncoding();
                byte[] key = UE.GetBytes(GlobalProperty.EncryptionPassword);

                //string cryptFile = outputFile;
                using (FileStream fsCrypt = new FileStream(EncryptedFile, FileMode.Create))
                {

                    RijndaelManaged RMCrypto = new RijndaelManaged();

                    using (CryptoStream cs = new CryptoStream(fsCrypt,
                                          RMCrypto.CreateEncryptor(key, key),
                                          CryptoStreamMode.Write))
                    {
                        for (int i = 0; i < inputdata.Length; i++)
                        {
                            cs.WriteByte((byte)inputdata[i]);
                        }
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool DecryptFile()
        {
            CryptoStream cs = null;
            FileStream fsCrypt = null;

            try
            {
                UnicodeEncoding UE = new UnicodeEncoding();
                byte[] key = UE.GetBytes(GlobalProperty.EncryptionPassword);

                using (fsCrypt = new FileStream(EncryptedFile, FileMode.Open))
                {
                    if (fsCrypt.Length == 0)
                    {
                        // this is first time. So no error msg.
                        mData.Message = null;
                        return false;
                    }

                    if (fsCrypt.Length> 0 && fsCrypt.Length < 8)
                    {
                        // sometime in error sitution. It has more than 0, but less than 8.
                        mData.Message = "Tyhjä encrypted.text tiedosto. Käynnistä verkkopalvelu uudelleen";
                        return false;
                    }

                    RijndaelManaged RMCrypto = new RijndaelManaged();

                    using (cs = new CryptoStream(fsCrypt,
                        RMCrypto.CreateDecryptor(key, key),
                        CryptoStreamMode.Read))
                    {
                        StreamReader reader = new StreamReader(cs);
                        string data = reader.ReadToEnd();
                        string[] splitData = data.Split(' ');
                        GlobalProperty.Username = splitData[0];
                        GlobalProperty.Password = splitData[1];

                        if (data.Length > 0)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string s = e.Message;
                return false;
            }
        }

        public ActionResult TKokous()
        {
            GlobalProperty.Calender = null;
            GC.Collect();
            GlobalProperty.Calender = new OutlookCalender();

            GlobalProperty.Calender.Initialize(GlobalProperty.Username, 
                                                 GlobalProperty.Password, 
                                                 GlobalProperty.AutodiscoverEmail);
            
            if (GlobalProperty.Calender.GetAllData())
            {
                if (GlobalProperty.Calender != null)
                {
                    if (GlobalProperty.Calender.TodayMeetingList.Count > 0)
                    {
                        GlobalProperty.Calender.RefreshSec = 0;

                        for(int i=0; i<GlobalProperty.Calender.TodayMeetingList.Count; i++)
                        {
                            GlobalProperty.Calender.RefreshSec += 10;
                        }
                        return View(GlobalProperty.Calender);
                    }
                    else
                    {
                        return RedirectToAction("HKokous", "Home");
                    }
                }
                else
                {
                    return RedirectToAction("Error", "Home");
                }
            }
            else
            {
                return RedirectToAction("Error", "Home");
            }
        }

        public ActionResult HKokous()
        {
            if (GlobalProperty.Calender != null)
            {
                if (GlobalProperty.Calender != null)
                {
                    if (GlobalProperty.Calender.TomorrowMeetingList.Count > 0)
                    {
                        GlobalProperty.Calender.RefreshSec = 0;

                        for (int i = 0; i < GlobalProperty.Calender.TomorrowMeetingList.Count; i++)
                        {
                            GlobalProperty.Calender.RefreshSec += 10;
                        }
                        return View(GlobalProperty.Calender);
                    }
                    else
                    {
                        return RedirectToAction("SPList", "Home");
                    }
                }
                else
                {
                    return RedirectToAction("Error", "Home");
                }
            }
            else 
            {
                return RedirectToAction("Error", "Home");
            }
        }

        public ActionResult Tervetuloa()
        {
            return View();
        }

        public ActionResult Error()
        {
            return View();
        }

        public ActionResult Submit(string user, string pass)
        {

            if (IsAuthenticated(user, pass))
            {
                string inputdata = user + " " + pass;
                // encrypt username and password
                if (!EncryptFile(inputdata))
                {
                    return RedirectToAction("Error", "Home");
                }
                GlobalProperty.Authenticated = true;
                if (Init(user, pass))
                {
                     return RedirectToAction("Tervetuloa", "Home");
                }
                else
                {
                    mData.Message = "Väärä käyttäjätunnus tai salasana. Yritä uudelleen";
                    GlobalProperty.Authenticated = false;
                    return RedirectToAction(actionName: "Index", routeValues: new { msg = mData.Message });
                }
            }
            else
            {
              mData.Message = "Väärä käyttäjätunnus tai salasana. Yritä uudelleen";
              GlobalProperty.Authenticated = false;
              return RedirectToAction(actionName: "Index", routeValues: new { msg = mData.Message });
            }
        }

        private bool Init(string userName, string password)
        {
            GlobalProperty.Username = userName;
            GlobalProperty.Password = password;
            if(String.IsNullOrEmpty(GlobalProperty.AutodiscoverEmail))
            {
                string email = GetEamil(userName);
                if (String.IsNullOrEmpty(email))
                {
                    return false;
                }
                else
                {
                    GlobalProperty.AutodiscoverEmail = email;
                }
            }
            return true;
        }
       
        public bool IsAuthenticated(string userName, string password)
        {
            string domainAndUsername = GlobalProperty.Domain + @"\" + userName;
            DirectoryEntry entry = new DirectoryEntry(GlobalProperty.LDAP, domainAndUsername, password);
            try
            {
                // Bind to the native AdsObject to force authentication.
                Object obj = entry.NativeObject;
                DirectorySearcher search = new DirectorySearcher(entry);
                search.Filter = "(SAMAccountName=" + userName + ")";
                search.PropertiesToLoad.Add("cn");
                SearchResult result = search.FindOne();
                if (null == result)
                {
                    return false;
                }
                // Update the new path to the user in the directory
                //Path = result.Path;
                //UserFullName = (String)result.Properties["cn"][0]; 
            }
            catch (Exception e)
            {
                string str = e.Message;
                return false;
            }
            return true;
        }

        private string GetEamil(string username)
        {
            try
            {
                DirectoryEntry entry = new DirectoryEntry();
                DirectorySearcher search = new DirectorySearcher(entry);
                // specify the search filter
                search.Filter = "(&(objectClass=user)(anr=" + username + "))";
                // specify which property values to return in the search
                search.PropertiesToLoad.Add("mail");          // smtp mail address
                //search.PropertiesToLoad.Add("givenName");   // first name
                //search.PropertiesToLoad.Add("sn");          // last name
                // perform the search
                search.Filter = String.Format("(sAMAccountName={0})", username);
                SearchResult result = search.FindOne();
                string email = result.Properties["mail"][0].ToString();
                //string givenname = result.Properties["givenName"][0].ToString();
                //string sn = result.Properties["sn"][0].ToString();
                return email;
            }
            catch (Exception e)
            {
                string str = e.Message;
                return null;
            }
        }

        public ActionResult SPList()
        {
            mData.ShareList.Clear();
            mData.RefreshSec = 0;
            if (GetInfoList())
            {
                if (mData.ShareList.Count > 0)
                {
                    return View(mData);
                }
                else
                {
                    return RedirectToAction("Tervetuloa", "Home");
                }
            }
            else
            {
                return RedirectToAction("Tervetuloa", "Home");
            }
        }
        private bool GetInfoList()
        {
            try
            {
                SpInfoList.EspooKaupunkitekniikkaDataContext DC = new SpInfoList.EspooKaupunkitekniikkaDataContext(
                    new Uri("http://tyotilat.espoo.fi/pali/espoo_kaupunkitekniikka/_vti_bin/ListData.svc"));
            
                DC.Credentials = new NetworkCredential(GlobalProperty.Username, GlobalProperty.Password, GlobalProperty.Domain);

                var source = DC.Ilmoitukset;

                if( source == null)
                {
                    return false;
                }

                foreach (var list in source)
                {
                    bool newBool = (list.InfoTV.HasValue) ? list.InfoTV.Value : false;

                    if ((bool)newBool)
                    {
                        ListItem item = new ListItem();
                        string htmlText = null;
                        if (ParseHtml(list.Teksti, out htmlText))
                        {
                            item.Teksti = htmlText;
                        }

                        item.Otsikko = list.Otsikko;
                        DateTime newVanhentuu = (list.Vanhentuu.HasValue) ? list.Vanhentuu.Value : DateTime.MinValue;
                        item.Vanhentuu = newVanhentuu.ToShortDateString();

                        Debug.WriteLine("Otsikko: " + list.Otsikko);
                        Debug.WriteLine("Teksti: " + list.Teksti);
                        Debug.WriteLine(" Vanhentuu: " + newVanhentuu.ToString());
                        mData.ShareList.Add(item);
                        mData.RefreshSec += 10;
                    }
                }

                if (mData.ShareList.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                string s = e.Message;
                return false;
            }
        }

        private bool ParseHtml(string inputStr, out string htmlText)
        {
            try
            {
                htmlText = null;
                HtmlDocument doc = new HtmlDocument();
                using (TextReader sr = new StringReader(inputStr))
                {
                    doc.Load(sr);

                    HtmlNode firstChild = null;
                    foreach (var node in doc.DocumentNode.SelectNodes("//div"))
                    {
                        HtmlNode lastNode2 = node.LastChild;
                        if (node == firstChild)
                        {

                            break;
                        }
                        if (node.HasChildNodes)
                        {
                            HtmlNode lastNode = node.LastChild;
                            var nodeCollection = node.ChildNodes;
                            foreach (var innterNode in nodeCollection)
                            {
                                string nodeText = innterNode.InnerText;

                                if (innterNode.Name == "html")
                                {
                                    Debug.WriteLine(nodeText);

                                    if (!String.IsNullOrWhiteSpace(nodeText))
                                    {
                                        htmlText += nodeText;
                                    }
                                }
                            }
                        }
                        else
                        {
                            htmlText += node.InnerText;
                            Debug.WriteLine(htmlText);
                        }
                        firstChild = node.FirstChild;
                    }
                    if (htmlText.Length > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                string s = e.Message;
                htmlText = null;
                return false;
            }
        }
    }
}
