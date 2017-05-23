using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using AuthorizeNet.Api.Controllers;
using AuthorizeNet.Api.Contracts.V1;
using AuthorizeNet.Api.Controllers.Bases;
using LouACH;
    namespace net.authorize.sample

{
    public class ChargeCreditCard
    {
        public static ANetApiResponse Run(String ApiLoginID, String ApiTransactionKey, decimal amount )
        {
            ApiOperationBase<ANetApiRequest, ANetApiResponse>.RunEnvironment = AuthorizeNet.Environment.SANDBOX;
            //ApiOperationBase<ANetApiRequest, ANetApiResponse>.RunEnvironment = AuthorizeNet.Environment.PRODUCTION;

            // define the merchant information (authentication / transaction id)
            ApiOperationBase<ANetApiRequest, ANetApiResponse>.MerchantAuthentication = new merchantAuthenticationType()
            {
                name = ApiLoginID,
                ItemElementName = ItemChoiceType.transactionKey,
                Item = ApiTransactionKey,
            };
            ;
 
            creditCardType creditCard;
            creditCard = LouACH.EventReceipt.creditCard;

            customerAddressType billingAddress;
            billingAddress = LouACH.EventReceipt.customerAddress;


             //standard api call to retrieve response
            var paymentType = new paymentType { Item = creditCard };

            var lineItems = new lineItemType[1];
            lineItems[0] = LouACH.EventReceipt.lineItems;
                   
            var transactionRequest = new transactionRequestType
            {
                transactionType = transactionTypeEnum.authCaptureTransaction.ToString(),    // charge the card

                amount = amount,
                payment = paymentType,
                billTo = billingAddress,
                lineItems = lineItems
            };
            
            var request = new createTransactionRequest { transactionRequest = transactionRequest };
            
            // instantiate the contoller that will call the service
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            var controller = new createTransactionController(request);
            controller.Execute();
            
            // get the response from the service (errors contained if any)
            var response = controller.GetApiResponse();

            //validate
            if (response != null)
            {
                if (response.messages.resultCode == messageTypeEnum.Ok)
                {
                    if(response.transactionResponse.messages != null)
                    {
                        LouACH.EventReceipt.TransactionMessage =  response.transactionResponse.messages[0].description + "(Transaction ID: " + response.transactionResponse.transId + ")";
 
                       //LouACH.EventMakePayment.sOutput = "Failed Transaction.";
                        if (response.transactionResponse.errors != null)
                        {
                            LouACH.EventReceipt.TransactionMessage = "Failed Transaction";
                        }
                    }
                }
                else
                {
                    if (response.transactionResponse != null && response.transactionResponse.errors != null)
                    {
                        LouACH.EventReceipt.TransactionMessage ="Transaction Error: " + response.transactionResponse.errors[0].errorText;
                        
                    }
                    else
                    {
                        LouACH.EventReceipt.TransactionMessage ="Transaction Error: " + response.transactionResponse.errors[0].errorText;

                    }
                }
            }
            else
            {
               LouACH.EventReceipt.TransactionMessage ="Transaction Error: No Response";

            }

            return response;
        }
    }
}
