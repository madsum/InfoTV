using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AspInfoTaulu.Models
{
    public class Data
    {
        public string Message { get; set; }
        public string Area { get; set; }
        public int RefreshSec { get; set; }
        public List<ListItem> ShareList { get; set; }

        public Data()
        {
            ShareList = new List<ListItem>();
        }
    }

    public class ListItem
    {
        public string Teksti { get; set; }
        public string Otsikko { get; set; }
        public string Vanhentuu { get; set; }

        public ListItem()
        { }
    }


}