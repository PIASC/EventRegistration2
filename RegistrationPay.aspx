<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="RegistrationPay.aspx.cs" Inherits="LouACH.RegistrationPay" %>

<!doctype html>
<!--[if lt IE 7]> <html class="ie6 oldie"> <![endif]-->
<!--[if IE 7]>    <html class="ie7 oldie"> <![endif]-->
<!--[if IE 8]>    <html class="ie8 oldie"> <![endif]-->
<!--[if gt IE 8]><!--> <html class="">
<!--<![endif]-->
<head runat="server">
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="PIASC's mission is simply to do everything possible to help our members succeed." />

<title>The Bob Lindgren Retirement Dinner</title>
<link href="../css/boilerplate.css" rel="stylesheet" type="text/css">
<link href="retirement.css" rel="stylesheet" type="text/css">

<!--[if lt IE 9]>
<script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
<![endif]-->
 <script> (function (d) {
     var config = {
         kitId: 'aem6egb',
         scriptTimeout: 3000
     },
     h = d.documentElement, t = setTimeout(function () { h.className = h.className.replace(/\bwf-loading\b/g, "") + " wf-inactive"; }, config.scriptTimeout), tk = d.createElement("script"), f = false, s = d.getElementsByTagName("script")[0], a; h.className += " wf-loading"; tk.src = '//use.typekit.net/' + config.kitId + '.js'; tk.async = true; tk.onload = tk.onreadystatechange = function () { a = this.readyState; if (f || a && a != "complete" && a != "loaded") return; f = true; clearTimeout(t); try { Typekit.load(config) } catch (e) { } }; s.parentNode.insertBefore(tk, s)
 })(document);
</script>
<script src="../js/respond.min.js"></script>

     <script type="text/javascript">
         function GoToNext() {
             fSummary.action = 'EventMakePayment.aspx';
             fSummary.submit();
             //window.location.href = 'RegistrationPay.aspx';
         }

</script>
</head>
<body>
    <form id="fSummary" runat="server">
       <header class="fluid orange">
<div class="gridContainer clearfix">
    <div id="div1" class="fluid"><img src="images/RetirementDinnerHeader.png" alt="banner"><br><br></div>
        <div id="top" class="fluid ">
        <ul class="fluid fluidlist nav">
            <li><a href="index.html">HOME</a></li>
        	<li id="rsvp"><a href="EventRegistration.aspx">RSVP</a></li>
            <li><a href="club.html">THE JONATHAN CLUB</a></li>
            <li><a href="hotels.html">HOTELS</a></li>
            <li><a href="bio.html">BIO</a></li>
            

        </ul>
        <div id="nav-toggle"><span></span></div>

        </div>
        
		<nav id="nav">
    <ul class="fluid fluidlist nav">
        	<li><a href="index.html">HOME</a></li>
        	<li><a href="EventRegistration.aspx">RSVP</a></li>
            <li><a href="club.html">THE JONATHAN CLUB</a></li>
            <li><a href="hotels.html">HOTELS</a></li>
            <li><a href="bio.html">BIO</a></li>
            

        </ul>
    </nav>
    
    <nav id="nav-mobile"></nav>
    
	</div>
        <div id="shadow">&nbsp;</div>

</header>
            	

<section class="fluid">

<div class="gridContainer clearfix">

<div id="main" class="fluid ">
    <div id="centered">
<div class="fluid" >
<p><strong>REGISTRATION SUMMARY</strong></p>
<div class="box1"  style="text-align:left">
   <p> <label><strong>Reservation For:</strong> <%=person.PersonfName%> <%=person.PersonlName%> <%=registration.LineItems%></label><br/>
    <%--<label><%=sMeal%> <%=sgMeal%></label><br />--%>
    <label><strong>Amount Due:</strong> <%=registration.Amount.ToString("C") %></label></p>
    <p><em>In lieu of gifts, please support one of Bob Lindgren's favorite foundations and causes:</em></p>
        <p><strong><a href="../raise/index.html" target="_blank">PIASC/RAISE Foundation</a></strong>   
       
            <asp:DropDownList ID="PIASCDonate" name="PIASCDonate" AutoPostBack="True"  OnSelectedIndexChanged="PIASCAddToTotal" runat="server">
              <asp:ListItem Text="0" Value="0" Selected="True"></asp:ListItem>
                        <asp:ListItem Text="$50" Value="50"></asp:ListItem>
                        <asp:ListItem Text="$100" Value="100"></asp:ListItem>
                        <asp:ListItem Text="$200" Value="200"></asp:ListItem>
                        <asp:ListItem Text="$300" Value="300"></asp:ListItem>
                        <asp:ListItem Text="$400" Value="400"></asp:ListItem>
                        <asp:ListItem Text="$500" Value="500"></asp:ListItem>
</asp:DropDownList></p> 
           <div class="donationList">
             <p>  The Foundation's goal is to foster graphic communications careers in primary, secondary and post-secondary educational institutions. RAISE is, and has been over the years, sustained by individual companies, industry employees, and other foundations.</p>
           </div>   
<p><strong><a href="http://printmuseum.org/" target="_blank">International Printing Museum</a></strong>
       
            <asp:DropDownList ID="IPMDonate" name="IPMDonate" AutoPostBack="True"  OnSelectedIndexChanged="PIASCAddToTotal" runat="server">
              <asp:ListItem Text="0" Value="0" Selected="True"></asp:ListItem>
                        <asp:ListItem Text="$50" Value="50"></asp:ListItem>
                        <asp:ListItem Text="$100" Value="100"></asp:ListItem>
                        <asp:ListItem Text="$200" Value="200"></asp:ListItem>
                        <asp:ListItem Text="$300" Value="300"></asp:ListItem>
                        <asp:ListItem Text="$400" Value="400"></asp:ListItem>
                        <asp:ListItem Text="$500" Value="500"></asp:ListItem>
           </asp:DropDownList></p> 
           <div class="donationList">
<p>This dynamic museum is devoted to bringing the history of books, printing, and the book arts to life for diverse audiences. The staff and volunteers make it their mission to take one of the world's most significant collections of antique printing machinery and interpret it for today's audiences through working demonstrations and theater presentations.</p>
           </div>   

       <p><strong><a href="http://www.printing.org/printpac-non-member" target="_blank">PrintPAC</a></strong>
       
            <asp:DropDownList ID="PPACDonate" name="PPACDonate" AutoPostBack="True"  OnSelectedIndexChanged="PIASCAddToTotal" runat="server">
              <asp:ListItem Text="0" Value="0" Selected="True"></asp:ListItem>
                        <asp:ListItem Text="$50" Value="50"></asp:ListItem>
                        <asp:ListItem Text="$100" Value="100"></asp:ListItem>
                        <asp:ListItem Text="$200" Value="200"></asp:ListItem>
                        <asp:ListItem Text="$300" Value="300"></asp:ListItem>
                        <asp:ListItem Text="$400" Value="400"></asp:ListItem>
                        <asp:ListItem Text="$500" Value="500"></asp:ListItem>
</asp:DropDownList></p> 
           <div class="donationList">
<p>The official political action committee of Printing Industries of America, is the industry's most important means of impacting public policy direction and debate pertaining to issues affecting printing and graphic communications companies. PIA member company CEOs, presidents, and executive management unite through PrintPAC to support federal candidates that are willing to defend and advance pro-print, pro-business legislation in Washington, DC.</p>

         

        <input type="hidden" id="fName" name="fName" value="<%=fName%>" /><input type="hidden" id="lName" name="lName" value="<%=lName%>" /><input type="hidden" id="gName"  name="gName" value="<%=gName%>" /><input type="hidden" id="Meal"  name="Meal" value="<%=Meal%>" /><input type="hidden" id="gMeal"  name="gMeal" value="<%=gMeal%>" /><input type="hidden" id="AmountDue"  name="AmountDue" value="<%=AmountDue%>" />
  <div>
    <br /><input id="ButtonComplete" type="button" class="Button" value="Complete Registration" onclick="GoToNext()" style="align-self:center"/>
 </div>
    </form>
     </div></div><br /><br /></section>
<div id="subFooter" class="fluid"><p>Bob Lindgren Retirement Dinner | Jonathan Club  | 545 S. Figueroa Street, Los Angeles, CA 90071 | 323.728.9500, Ext. 229</p></div>
<script src="../js/vendor/jquery-1.11.1.min.js"></script>
<script src="../js/vendor/modernizr-2.8.3.min.js"></script>
<script src="../js/plugins.js"></script>
<script src="../js/main.js"></script>

</body>
</html>
