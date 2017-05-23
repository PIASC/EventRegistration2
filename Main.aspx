<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Main.aspx.cs" Inherits="LouACH.Main" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style>
        p
        {
            font-style:italic;
            text-align:right;

        }
        .dAccount
        {
            float:left;
            left:100px;
            top:100px;
            width:350px;
            padding:10px;
            border:2px solid Black;

        }
        .TestFloaT
        {
            float:left;
            width:200px;
            height:300px;
           
        }
        .PayResponse
        {
            
            width:250px;
            height:100px;
            padding:10px;
            border:2px solid Black;
            font-style:italic;
            font-size:x-small;
        }
        .PayResponse.Before
        {
          visibility:hidden;
        }
         .PayResponse.After
        {
          visibility:visible;
        }
        h1
        {
            font-size:medium;
            text-align:center;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="TestFloaT">

        </div>
    <div class="dAccount">
        <p>
            <asp:DropDownList ID="ddlEmployers" name="ddlEmployers" AutoPostBack="True"
                    OnSelectedIndexChanged="Selection_Change" public="True" runat="server">

            </asp:DropDownList>
        </p>
        <p>
        Company Name:
        <asp:TextBox ID="tCompanyName"  runat="server" placeholder=" Name of Company"/>
        </p>
        <p>
            Bank Name:
         <asp:TextBox ID="tBankName"  runat="server"  placeholder="Name of Bank"/>
        </p>
         <p>
             Routing Number:
        <asp:TextBox ID="tRouting"  runat="server" placeholder="9 digits required"/>
        </p>
        <p>
            Account Name:
        <asp:TextBox ID="tAccount"  runat="server" />
        </p>
        <p>
            Current Amount Due:
         <asp:TextBox ID="tAmount" runat="server"  placeholder="Amount Due" />
        </p>
        <p>
        <%--<asp:Button ID="Button1" runat="server" Text="Pay Amount" OnClick="Button1_click" />--%>
        <asp:Button ID="bMakePayment" runat="server" Text="Pay Amount" OnClick="bMakePayment_click" />
        </p>
        <p>
            <asp:Button ID="bMakeProfile" runat="server" Text="Make Profile" OnClick="bMakeProfile_click" />
        </p>
        <p>
            <asp:Button ID="bGetProfile" runat="server" Text="Get Profile" OnClick="bGetProfile_click" />
        </p>
        <p>
            <asp:Button ID="bChargeProfile" runat="server" Text="Pay Profile" OnClick="bChargeProfile_click" />
        </p>   
        <p>
            <asp:Button ID="bChargeCC" runat="server" Text="Charge CC" OnClick="bChargeCC_click" />
        </p>        
        <div class="PayResponse" id="PayResponse" runat="server" >
        <h1>RESPONSE</h1>
            <p><%=theOutput %></p>            
        </div>
    </div>

    </form>
</body>
</html>
