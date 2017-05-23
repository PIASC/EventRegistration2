using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LouACH.Events
{
    public class EventTransaction
    {
        public int TransactionID;
        public int RegistrationID;
        public string TransactionCode;
        public decimal AmountPaid;
        public string LineItem;
        public int AccountID;
    }
} 