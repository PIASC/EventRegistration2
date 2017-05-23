using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace LouACH.Events
{
    public class CardConnectRestClient
    {
    public  string url;
	public string user;
    public string password;

      public static void authorizeTransaction(Newtonsoft.Json.Linq.JObject request)
        {
    
          //return self::send($this->ENDPOINT_AUTH, $this->OP_PUT, $request);
        }
    
    }
    //     public static authorizeTransaction(request)) 
    //{
    //    return self::send($this->ENDPOINT_AUTH, $this->OP_PUT, $request);
    //}
}
