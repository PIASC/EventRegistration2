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
    
    public partial class Main : System.Web.UI.Page
    {
        public static string EmployerID = "1"; 
        public static decimal Amount;
        public static string theOutput = "";
        public static string apiLoginId = Constants.API_LOGIN_ID;
        public static string transactionKey = Constants.TRANSACTION_KEY;
        public static string TransactionID = Constants.TRANSACTION_ID;
        public static bankAccountType bankAccount= LouACH.DataBaseTransactions.DataBase.ReadBankData();
        protected void Page_Load(object sender, EventArgs e)
        {
            //    string apiLoginId = SampleCodeTest.Constants.API_LOGIN_ID;
            //    string transactionKey = SampleCodeTest.Constants.TRANSACTION_KEY;
            ;         //    string TransactionID = SampleCodeTest.Constants.TRANSACTION_ID;
            //    string payerID = SampleCodeTest.Constants.PAYER_ID;
            //    decimal Amount = 100.00M;
            //    net.authorize.sample.DebitBankAccount.Run(apiLoginId, transactionKey, Amount);
            //Amount = 1000.00M;
            //
            //bankAccountType bankAccount;
            //Response.Redirect("EventRegistration.aspx");
            bankAccount = LouACH.DataBaseTransactions.DataBase.ReadBankData();
            tCompanyName.Text = bankAccount.nameOnAccount;
            tRouting.Text = bankAccount.routingNumber;
            tAccount.Text = bankAccount.accountNumber;
            tBankName.Text = bankAccount.bankName;
            Amount = LouACH.DataBaseTransactions.DataBase.AmountDue();
            tAmount.Text = System.Convert.ToString(Amount);
            PayResponse.Attributes.Add("Class", "PayResponse Before");
            if (!this.IsPostBack)
            {
 
                DataTable table = LouACH.DataBaseTransactions.DataBase.getEmployersDropDowndata();
                ddlEmployers.Items.Clear();
                ddlEmployers.DataValueField = "EMPLOYERID";
                ddlEmployers.DataTextField = "ACCTNAME";
                ddlEmployers.DataSource = table;
                ddlEmployers.DataBind();
            }

        }
        public void bMakePayment_click(object sender, EventArgs e)
        {
            //string apiLoginId = SampleCodeTest.Constants.API_LOGIN_ID;
            //string transactionKey = SampleCodeTest.Constants.TRANSACTION_KEY;
            //string TransactionID = SampleCodeTest.Constants.TRANSACTION_ID;
            //string payerID = SampleCodeTest.Constants.PAYER_ID;
            //string apiLoginId = Constants.API_LOGIN_ID;
            //string transactionKey = Constants.TRANSACTION_KEY;
            //string TransactionID = Constants.TRANSACTION_ID;
            //string payerID = Constants.PAYER_ID;           //decimal Amount = 100.00M;

            // Get Value
            //Amount = LouACH.DataBaseTransactions.DataBase.AmountDue();
            Amount = System.Convert.ToDecimal(tAmount.Text);
            //var BankData = new bankAccountType  
            //        {
            //            accountNumber = "00000000",
            //            routingNumber = "00000000",
            //            echeckType = echeckTypeEnum.WEB,   // change based on how you take the payment (web, telephone, etc)
            //            nameOnAccount = "default"
            //        };
            //ReadData("Data Source=localhost:1521/XE;User Id=HSG_USER1;Password=Hsg45900;");

            var response = net.authorize.sample.DebitBankAccount.Run(apiLoginId, transactionKey, EmployerID, Amount);
            //PayResponse.Attributes.Clear("PayResponse Before");
            PayResponse.Attributes.Add("Class", "PayResponse After");
            //PayResponse.Visible = true;
            //ReadData("Data Source=localhost:1521/XE;User Id=HSG_USER1;Password=Hsg45900;");
        }
        public void bChargeProfile_click(object sender, EventArgs e)
        {
            Amount = System.Convert.ToDecimal(tAmount.Text);

            var response = net.authorize.sample.ChargeCustomerProfile.Run(apiLoginId, transactionKey, "1810413571", "1805162419",Amount);
            PayResponse.Attributes.Add("Class", "PayResponse After");
        }

        public void bMakeProfile_click(object sender, EventArgs e)
        {
            //var response = net.authorize.sample.GetCustomerPaymentProfileList.Run(apiLoginId, transactionKey);
            //var response = net.authorize.sample.GetCustomerPaymentProfile.Run(apiLoginId, transactionKey,"1810413571","1805162419");
            var response = net.authorize.sample.CreateCustomerProfile.Run(apiLoginId, transactionKey, EmployerID);
            //PayResponse.Attributes.Clear("PayResponse Before");
            PayResponse.Attributes.Add("Class", "PayResponse After");
            //PayResponse.Visible = true;
            //ReadData("Data Source=localhost:1521/XE;User Id=HSG_USER1;Password=Hsg45900;");
        }
        public void bGetProfile_click(object sender, EventArgs e)
        {
            //var response = net.authorize.sample.GetCustomerPaymentProfileList.Run(apiLoginId, transactionKey);
            var response = net.authorize.sample.GetCustomerPaymentProfile.Run(apiLoginId, transactionKey,"1810413571","1805162419");
            //var response = net.authorize.sample.CreateCustomerProfile.Run(apiLoginId, transactionKey, EmployerID);
            //PayResponse.Attributes.Clear("PayResponse Before");
            PayResponse.Attributes.Add("Class", "PayResponse After");
            //PayResponse.Visible = true;
            //ReadData("Data Source=localhost:1521/XE;User Id=HSG_USER1;Password=Hsg45900;");
        }

        public void Selection_Change(object sender, EventArgs e)
        {
            EmployerID = ddlEmployers.SelectedValue;
            // ddlEmployers.SelectedIndex = ddlEmployers.SelectedValue;
            //ddlEmployers.Items.Clear();
            bankAccount = LouACH.DataBaseTransactions.DataBase.ReadBankData();
            tCompanyName.Text = bankAccount.nameOnAccount;
            tRouting.Text = bankAccount.routingNumber;
            tAccount.Text = bankAccount.accountNumber;
            tBankName.Text = bankAccount.bankName;
            Amount = LouACH.DataBaseTransactions.DataBase.AmountDue();
            tAmount.Text = System.Convert.ToString(Amount);
            PayResponse.Attributes.Add("Class", "PayResponse Before");
        }
        public void bChargeCC_click(object sender, EventArgs e)
        {
            var response = net.authorize.sample.ChargeCreditCard.Run(apiLoginId, transactionKey, Amount);

        }
 
        //private static void ReadData(string connectionString)
        //{
        //    string queryString = "select * from EMPLOYER where EMPLOYERID=1";
        //    using (OracleConnection connection = new OracleConnection(connectionString))
        //    {
        //        OracleCommand command = new OracleCommand(queryString, connection);
        //        connection.Open();
        //        OracleDataReader reader;
        //        reader = command.ExecuteReader();

        //        // Always call Read before accessing data.
        //        while (reader.Read())
        //        {
        //            theOutput = theOutput + "<p/>" + reader.GetString(7);
        //            //Console.WriteLine(reader.GetInt32(0) + ", " + reader.GetString(1));
        //            var bankAccount = new bankAccountType
        //            {
        //                accountNumber = Convert.ToString(reader.GetInt32(4)),
        //                routingNumber = Convert.ToString(reader.GetValue(6)),
        //                echeckType = echeckTypeEnum.WEB,   // change based on how you take the payment (web, telephone, etc)
        //                nameOnAccount = reader.GetString(7)
        //            };
        //         //return bankAccount;
        //        }
               
        //        // Always call Close when done reading.
        //        reader.Close();
        //    }
        //}

}
}

