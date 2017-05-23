using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace LouACH
{
    public partial class RegistrationPay : System.Web.UI.Page
    {
        public static string fName = "";
        public static string lName = "";
        public static string gName = "";
        public static string sgName = "";
        public static string Meal = "";
        public static string sMeal = "";
        public static string gMeal = "";
        public static string sgMeal = "";
        public static decimal AmountDue = 0.00m;
        public static decimal gAmount = 0.00m;
        public static decimal Amount1 = 0.00m;
        public static decimal Amount2 = 0.00m;
        public static decimal Amount3 = 0.00m;
        public static LouACH.Events.Person person;
        public static LouACH.Events.Registration registration;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                person = new Events.Person
                {
                    PersonfName = Request.Form["txtFName"],
                    PersonlName = Request.Form["txtLName"],
                    PersonEMail = Request.Form["txtEMail"],
                    PersonCompany = Request.Form["txtCoName"]
                };
                person.PersonID = Convert.ToInt32(LouACH.DataBaseTransactions.DataBase.GetPersonID(person));
                registration = new Events.Registration
                {
                    EventID = 1,
                    PersonID = person.PersonID,
                    Amount = 0.00m,
                    LineItems = "Meal:" + Request.Form["rMeal"]
                };
                Meal = Request.Form["rMeal"];
                registration.RegistrationID = Convert.ToInt32(LouACH.DataBaseTransactions.DataBase.GetRegistrationID(registration));
                Session.Add("RegistrationID",registration.RegistrationID) ;
                //if (Request.Form["txtGuestName"] != "")
                if (Request.Form["CheckGuest"] == "1")
                    {
                    registration.LineItems = registration.LineItems + ",GuestName:" + Request.Form["txtGuestName"] + ",Guest Meal:" + Request.Form["gMeal"];
                    gName = Request.Form["txtGuestName"];
                    gMeal = Request.Form["gMeal"];
                    sgName = " and " + Request.Form["txtGuestName"];
                    sgMeal = " and " + Request.Form["gMeal"];
                    gAmount = 200.00m;
                    AmountDue = 200.00m;
                    registration.Amount = AmountDue;
                     }
                else
                {
                    gAmount = 0.00m;
                    AmountDue = 0.00m;
                    registration.Amount = AmountDue;               
                
                };

            }
        }



        protected void PIASCAddToTotal(object sender, EventArgs e)
        {
            Amount1 = System.Convert.ToDecimal(PIASCDonate.SelectedItem.Value);
            Amount2 = System.Convert.ToDecimal(IPMDonate.SelectedItem.Value);
            Amount3 = System.Convert.ToDecimal(PPACDonate.SelectedItem.Value);
            AmountDue = gAmount + Amount1;
            AmountDue = AmountDue + Amount2;
            AmountDue = AmountDue + Amount3;
            registration.Amount = AmountDue;
        }
    }
}
