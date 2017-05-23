using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace LouACH
{
    public partial class EventRegistration : System.Web.UI.Page
    {
        public static LouACH.Events.Event thisEvent = LouACH.DataBaseTransactions.DataBase.GetEvent("1");
        
        protected void Page_Load(object sender, EventArgs e)
        {
            string[] a = { "a", "b", "c" };
            //CardConnectRestClientExample.CardConnectRestClientExample.Main(a);
        }
    }
}