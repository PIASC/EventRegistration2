using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using AuthorizeNet.Api.Contracts.V1;
using net.authorize.sample;

namespace LouACH
{
    public partial class EventMakePayment : System.Web.UI.Page
    {
        public static creditCardType creditCard;
        public static customerAddressType customerAddress;
        public static string fName = "";
        public static string lName = "";
        public static string gName = "";
        public static string Meal = "";
        public static string gMeal = "";
        public static string sAmountDue = "0.00";
        public static string PIASC = "0.00";
        public static string IPM = "0.00";
        public static string PPAC = "0.00";
        public static decimal AmountDue = 0.00m;
        public static LouACH.Events.EventTransaction Transaction;
        public static List<Events.EventTransaction> allTransactions;
        protected void Page_Load(object sender, EventArgs e)
        {

            fName = LouACH.RegistrationPay.person.PersonfName;
            lName = LouACH.RegistrationPay.person.PersonlName;
            Meal = Request.Form["Meal"];    //LouACH.RegistrationPay.registration.LineItems;

            AmountDue = LouACH.RegistrationPay.registration.Amount;
            if (AmountDue == 0.00m)
            {
                Server.Transfer("EventReceiptNoCharge.aspx?fName=" + fName + "&lName=" + lName + "&Meal=" + Meal, true);
            }
            {
                sAmountDue = AmountDue.ToString("$#,##0.00");
            }
            
            List<Events.EventTransaction> allTransactions = new List<Events.EventTransaction> ();

                if (Request.Form["gName"] != "")
                {
                    Transaction = new Events.EventTransaction
                    {
                        RegistrationID = Convert.ToInt32(Session["RegistrationID"]),
                        AmountPaid = 200.00m,
                        LineItem = "Guest Meal",   //" + Request.Form["gName"] + " meal(" + Request.Form["gMeal"] + ")",
                        AccountID=1
                    };
                    Transaction.TransactionID = Convert.ToInt32(LouACH.DataBaseTransactions.DataBase.SaveTransaction(Transaction));
                    //allTransactions.Add(Transaction1);
                    
                    gName = Request.Form["gName"];
                    gMeal = Request.Form["gMeal"];

                };
                
                if (Request.Form["PIASCDonate"] != "0")
                {  
                    Transaction = new Events.EventTransaction
                    {
                        RegistrationID = Convert.ToInt32(Session["RegistrationID"]),
                        AmountPaid= System.Convert.ToDecimal(Request.Form["PIASCDonate"]),
                        LineItem = "PIASC Donation",
                        AccountID = 2
                    };
                    Transaction.TransactionID = Convert.ToInt32(LouACH.DataBaseTransactions.DataBase.SaveTransaction(Transaction));


                };
                if (Request.Form["IPMDonate"] != "0")
                {  
                    Transaction = new Events.EventTransaction
                    {
                        RegistrationID = Convert.ToInt32(Session["RegistrationID"]),
                        AmountPaid= System.Convert.ToDecimal(Request.Form["IPMDonate"]),
                        LineItem = "IPM Donatation",
                        AccountID = 3
                    };
                    Transaction.TransactionID = Convert.ToInt32(LouACH.DataBaseTransactions.DataBase.SaveTransaction(Transaction));


                };
                 if (Request.Form["PPACDonate"] != "0")
                {  
                    Transaction = new Events.EventTransaction
                    {
                        RegistrationID = Convert.ToInt32(Session["RegistrationID"]),
                        AmountPaid= System.Convert.ToDecimal(Request.Form["PPACDonate"]),
                        LineItem = "PPAC Donatation",
                        AccountID = 4
                    };
                    Transaction.TransactionID = Convert.ToInt32(LouACH.DataBaseTransactions.DataBase.SaveTransaction(Transaction));


                };

             }

        }
    }
