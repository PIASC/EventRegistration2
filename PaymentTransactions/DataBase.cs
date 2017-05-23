using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using AuthorizeNet.Api.Controllers;
using AuthorizeNet.Api.Contracts.V1;
using AuthorizeNet.Api.Controllers.Bases;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
using LouACH;
using net.authorize.sample;

namespace LouACH.DataBaseTransactions
{
    public class DataBase
    {
        //public static string EmployerID = "2";
        public static string connectionString = Constants.CONNECTION_STRING_Local;
        public static bankAccountType ReadBankData()
        {
            string queryString = "select A.*,B.BANKNAME from EMPLOYER A join ROUTINGNOS B on A.TRANSITNO=B.BANKROUTINGNUMBER where EMPLOYERID=" + LouACH.Main.EmployerID;
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                OracleCommand command = new OracleCommand(queryString, connection);
                connection.Open();
                OracleDataReader reader;
                reader = command.ExecuteReader();

                // Always call Read before accessing data.
                //while (reader.Read())
                //{
                //    theOutput = theOutput + "<p/>" + reader.GetString(7);
                //    //Console.WriteLine(reader.GetInt32(0) + ", " + reader.GetString(1));
                reader.Read();
                var bankAccount = new bankAccountType
                {
                    routingNumber = Convert.ToString(reader.GetInt32(4)),
                    accountNumber = Convert.ToString(reader.GetValue(6)),
                    echeckType = echeckTypeEnum.WEB,   // change based on how you take the payment (web, telephone, etc)
                    nameOnAccount = reader.GetString(7),
                    bankName=reader.GetString(8)
                };
                reader.Close();
                return bankAccount;
                //}
                //reader.Read();
                //return reader;
                // Always call Close when done reading.
                //reader.Close();
            };
        }

        public static string WriteTransactionData(string TransID, decimal Amount,bankAccountType bankAccount)
        {
            string queryString = "insert into ACHITEM (TRANSITNO,BANKACCT,AMOUNT,ACCTNAME,TRANSDATETIME,TRANSID) values(:TS_NO,:BK_ACC,:amount,:AC_NAM,CURRENT_TIMESTAMP,:TS_ID)";
            //string queryString = "insert into ACHITEM (TRANSITNO,BANKACCT,AMOUNT,ACCTNAME) values(:TS_NO,:BK_ACC,:amount,:AC_NAM)";
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                command.Parameters.Add("TS_NO", bankAccount.routingNumber);
                command.Parameters.Add("BK_ACC", bankAccount.accountNumber);
                command.Parameters.Add("amount", Amount);
                command.Parameters.Add("AC_NAM", bankAccount.nameOnAccount);
                command.Parameters.Add("TS_ID", TransID);
                //command.Parameters.Add("TS_TIM", DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss"));
                //command.Parameters.Add("TS_TIM", DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss"));
                command.Connection.Open();
                command.ExecuteNonQuery();
                command.Connection.Close();
            }
            return "Success";
        }
        public static Decimal AmountDue()
        {
            string queryString = "select * from ACCOUNTS_DUE where EMPLOYERID=" + LouACH.Main.EmployerID;
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                OracleCommand command = new OracleCommand(queryString, connection);
                connection.Open();
                OracleDataReader reader;
                reader = command.ExecuteReader();

                // Always call Read before accessing data.
                //while (reader.Read())
                //{
                //    theOutput = theOutput + "<p/>" + reader.GetString(7);
                //    //Console.WriteLine(reader.GetInt32(0) + ", " + reader.GetString(1));
                reader.Read();
                return reader.GetDecimal(2);
            };
        }

        //public static void getEmployersDropDowndataX()
        //{
        //    string queryString = "select A.* from EMPLOYER A  where EMPLOYERID=" + EmployerID;
        //    using (OracleConnection connection = new OracleConnection(connectionString))
        //    {
        //        OracleCommand command = new OracleCommand(queryString, connection);
        //        connection.Open();
        //        OracleDataReader EmployerData;
        //        EmployerData = command.ExecuteReader();
        //        if (EmployerData.HasRows)
        //        {
        //            LouACH.Main.ddlEmployers.DataSource = EmployerData;
        //            LouACH.Main.ddlEmployers.DataValueField = "EMPLOYERID";
        //            LouACH.Main.ddlEmployers.DataTextField = "ACCTNAME";
        //            LouACH.Main.ddlEmployers.DataBind();
        //        }
        //        else
        //        {
        //            //MessageBox.Show("No is found");
        //            //CloseConnection = new connection();
        //            //CloseConnection.closeConnection(); // close the connection 
        //        }
        //    }
        //}

        public static DataTable getEmployersDropDowndata()
        {
            string cmdstr = "select A.* from EMPLOYER A order by EMPLOYERID";
            
            using (OracleConnection constr = new OracleConnection(connectionString))
            {
                OracleDataAdapter adapter = new OracleDataAdapter(cmdstr, constr);
                // Create the builder for the adapter to automatically generate
                // the Command when needed
                //OracleCommandBuilder builder = new OracleCommandBuilder(adapter);

                // Create and fill the DataSet using the EMP
                DataSet dataset = new DataSet();
                adapter.Fill(dataset, "EMP");

                // Get the EMP table from the dataset
               DataTable table = dataset.Tables["EMP"];
               return table;
                //LouACH.Main.ddlEmployers.DataValueField = "EMPLOYERID";
                //LouACH.Main.ddlEmployers.DataTextField = "ACCTNAME";
                //LouACH.Main.ddlEmployers.DataSource = table;

                //LouACH.Main.ddlEmployers.DataBind();
                
                
                
                //OracleCommand command = new OracleCommand(queryString, connection);
                //connection.Open();
                //OracleDataReader EmployerData;
                //EmployerData = command.ExecuteReader();
                //if (EmployerData.HasRows)
                //{
                //    LouACH.Main.ddlEmployers.DataSource = EmployerData;
                //    LouACH.Main.ddlEmployers.DataValueField = "EMPLOYERID";
                //    LouACH.Main.ddlEmployers.DataTextField = "ACCTNAME";
                //    LouACH.Main.ddlEmployers.DataBind();
                //}
                //else
                //{
                //    //MessageBox.Show("No is found");
                //    //CloseConnection = new connection();
                //    //CloseConnection.closeConnection(); // close the connection 
                //}
            }
        }
   
        
        
        
        
        
        //        //OracleConnection conn = new OracleConnection();conn.ConnectionString = ConfigurationManager.ConnectionStrings["ConnetionString"].ConnectionString;
                //    OracleConnection conn = new OracleConnection(connectionString)    
                //    OracleCommand cmd = new OracleCommand("SELECT EMPLOYEE_ID ID, EMPLOYEE_NAME VALUE FROM EMPLOYEE MASTER",conn);
                //        OracleDataAdapter oda = new OracleDataAdapter(cmd);
                //        DataSet ds = new DataSet();
                //        oda.Fill(ds);
                //        Dropdown.DataSource = ds;
                //        Dropdown.DataTextField = "VALUE";
                //        Dropdown.DataValueField = "ID";
                //        Dropdown.DataBind();
        
                //}
               
                //public DataTable DataTableQuery(string sql)
                //{
                //    DataTable table = new DataTable();
                //    try
                //    {
                //        using (OracleConnection connection = new OracleConnection(_connectString))
                //        {
                //            connection.Open();
                //            using (OracleCommand oracleCommand = new OracleCommand())
                //            {
                //                oracleCommand.CommandText = Regex.Replace(sql, @"\s+", " ");
                //                oracleCommand.Connection = connection;
                //                using (
                //                    OracleDataReader dataReader = oracleCommand.ExecuteReader(CommandBehavior.CloseConnection))
                //                {
                //                    table.Load(dataReader);
                //                }
                //            }
                //        }
                //    }
                //    catch (Exception e)
                //    {
                //        Console.WriteLine(e);
                //    }
                //    return table;
                //}
                //protected virtual DataTable ConnectAndQuery()
                //{
                //    OracleConnection conn = new OracleConnection();
                //    conn.ConnectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
                //    conn.Open();
                //    OracleCommand command = conn.CreateCommand();
                //    string sql = "SELECT EMPLOYEE_ID ID, EMPLOYEE_NAME VALUE FROM EMPLOYEE MASTER";
                //    command.CommandText = sql;
                //    OracleDataReader reader = command.ExecuteReader();
                //    DataTable dtusertables = new DataTable();
                //    dtusertables.Load(reader);
                //    if (null != reader && !reader.IsClosed)
                //    {
                //        reader.Close();
                //    }
                //    return dtusertables;
                //}
        public static LouACH.Events.Event GetEvent(string sEventID)
        {
            string queryString = "select * From EVENTS where EventID=" + sEventID;
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                OracleCommand command = new OracleCommand(queryString, connection);
                connection.Open();
                OracleDataReader reader;
                reader = command.ExecuteReader();

                reader.Read();
                var thisEvent = new LouACH.Events.Event
                {
 		        EventName= reader.GetString(1),
		        EventVenueName= reader.GetString(3),
		        //EventVenueAddress= reader.GetString(7);
		        //EventVenueCity= reader.GetString(7);
		        //EventVenueState= reader.GetString(7);
		        //EventVenueZip= reader.GetString(7);
                EventVenueDateTime = reader.GetDateTime(2)
                  };
                reader.Close();
                return thisEvent;
            };
        }
        public static string WriteEventTransactionDataNoCharge(LouACH.Events.Registration registration)
        {
            string queryString = "insert into EVENT_REGISTRATION (RegistrationID,EventID,PersonID,RegistrationDate,ITEM) values(1,:EV_ID,:P_ID,CURRENT_TIMESTAMP,:P_ITM)";

            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                //command.Parameters.Add("REG_NO", EventTransaction.routingNumber);
                //command.Parameters.Add("EV_ID", EventTransaction.EventID);
                //command.Parameters.Add("P_ID", EventTransaction.AttendeeID);
                //command.Parameters.Add("AMT", EventTransaction.AmountPaid);
                //command.Parameters.Add("REG_NO", 5);
                command.Parameters.Add("EV_ID", registration.EventID);
                command.Parameters.Add("P_ID", registration.PersonID);
                command.Parameters.Add("P_ITM", registration.LineItems);
                command.Connection.Open();
                command.ExecuteNonQuery();
                command.Connection.Close();
            }
            return "Success";
        }
        public static string GetPersonID(LouACH.Events.Person  person)
        {
            string queryString = "insert into PERSON(PersonID,FirstName,LastName,EMail,Company) values(1,:P_FN,:P_LN,:P_EM,:P_CO) RETURNING PersonID INTO :id" ;
            string Response = "";
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                command.Parameters.Add("P_FN", person.PersonfName);
                command.Parameters.Add("P_LN", person.PersonlName);
                command.Parameters.Add("P_EM", person.PersonEMail);
                command.Parameters.Add("P_CO", person.PersonCompany);
                command.Parameters.Add(new OracleParameter
                    (
                        ":id",
                        OracleDbType.Int32,
                        ParameterDirection.Output
                    ));
                command.Connection.Open();
                command.ExecuteNonQuery();
                Response=command.Parameters[":id"].Value.ToString();
                command.Connection.Close();
            }
            return Response;
        }
        public static string GetRegistrationID(LouACH.Events.Registration registration)
        {
            string queryString = "insert into EVENT_REGISTRATION values(1,:R_EVID,:R_PID,CURRENT_TIMESTAMP,:R_AMT,:R_ITM) RETURNING RegistrationID INTO :id";
            string Response = "";
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                command.Parameters.Add("R_EVID", registration.EventID);
                command.Parameters.Add("R_PID", registration.PersonID);
                command.Parameters.Add("R_AMT", registration.Amount);
                command.Parameters.Add("R_ITM", registration.LineItems);
                command.Parameters.Add(new OracleParameter
                    (
                        ":id",
                        OracleDbType.Int32,
                        ParameterDirection.Output
                    ));
                command.Connection.Open();
                command.ExecuteNonQuery();
                Response = command.Parameters[":id"].Value.ToString();
                command.Connection.Close();
            }
            return Response;
        }
        public static string SaveTransaction(LouACH.Events.EventTransaction transaction)
        {
            string queryString = "insert into EVENT_TRANSACTIONS values(1,:T_RGID,CURRENT_TIMESTAMP,:T_CD,:T_ST,:T_LN,:T_AC) RETURNING TransactionID INTO :id";
            string Response = "";
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                command.Parameters.Add("T_RGID", transaction.RegistrationID);
                command.Parameters.Add("T_CD", transaction.TransactionCode);
                command.Parameters.Add("T_ST", transaction.AmountPaid);
                command.Parameters.Add("T_LN", transaction.LineItem);
                command.Parameters.Add("T_AC", transaction.AccountID);
                command.Parameters.Add(new OracleParameter
                    (
                        ":id",
                        OracleDbType.Int32,
                        ParameterDirection.Output
                    ));
                command.Connection.Open();
                command.ExecuteNonQuery();
                Response = command.Parameters[":id"].Value.ToString();
                command.Connection.Close();
            }
            return Response;
        }
         
        public static string UpdateTransactions(string TransactionID, string TransactionCode)
        {
            string queryString = "UPDATE EVENT_TRANSACTIONS set TransactionCode =:T_CD where TransactionID =  :keyValue";
            string Response = "";
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                //command.Parameters.Add("T_CD", transaction.TransactionCode);
                //command.Parameters.Add("keyValue", transaction.TransactionID);
                command.Parameters.Add("T_CD", TransactionCode);
                command.Parameters.Add("keyValue", TransactionID);
                command.Connection.Open();
                command.ExecuteNonQuery();
                command.Connection.Close();
            }
            return Response;
        }
        public static OracleDataReader GetTransactions()
        {
            OracleDataReader dataReader;
            string queryString = "Select * FROM EVENT_TRANSACTIONS where RegistrationID = " + LouACH.RegistrationPay.registration.RegistrationID;
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                command.Connection.Open();
                dataReader = command.ExecuteReader(CommandBehavior.CloseConnection);
                return dataReader;
                //command.Connection.Close();
            }

        }
        public static LouACH.Events.Registration GetRegistration(string sRegistrationID)
        {
            OracleDataReader dataReader;
            string queryString = "Select * FROM EVENT_REGISTRATION where RegistrationID = " + sRegistrationID;
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                command.Connection.Open();
                dataReader = command.ExecuteReader(CommandBehavior.CloseConnection);
                LouACH.Events.Registration registration = new LouACH.Events.Registration
                {
                            RegistrationID=dataReader.GetInt32(0),
                            EventID=dataReader.GetInt32(1),
                            PersonID=dataReader.GetInt32(2),
                            Amount=dataReader.GetDecimal(4)
                };

                command.Connection.Close();
                return registration;
               
            }

        }

        public static LouACH.Events.Person GetPerson(string sRegistrationID)
        {
            OracleDataReader dataReader;
            string queryString = "select * from Person A join EVENT_REGISTRATION B on A.PERSONID=B.PERSONID where B.REGISTRATIONID=" + sRegistrationID;
            using (OracleConnection connection = new OracleConnection(connectionString))
            using (OracleCommand command = new OracleCommand(queryString, connection))
            {
                command.Connection.Open();
                dataReader = command.ExecuteReader();
                dataReader.Read();
                LouACH.Events.Person person = new LouACH.Events.Person
                {
                    PersonID = dataReader.GetInt32(0),
                    PersonfName = dataReader.GetString(2),
                    PersonlName = dataReader.GetString(1),
                    //PersonAddress = dataReader.GetString(4),
                    //PersonCity = dataReader.GetString(5),
                    //PersonState = dataReader.GetString(6),
                    //PersonZip = dataReader.GetString(7),
                    //PersonPhone="",
                    PersonEMail = dataReader.GetString(3),
                    PersonCompany=dataReader.GetString(8)
                };

                command.Connection.Close();
                return person;

            }

        }
    }
}