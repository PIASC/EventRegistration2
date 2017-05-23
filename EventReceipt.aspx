<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EventReceipt.aspx.cs" Inherits="LouACH.EventReceipt" %>

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
 <script> (function(d) {
    var config = {
      kitId: 'aem6egb',
      scriptTimeout: 3000
    },
    h=d.documentElement,t=setTimeout(function(){h.className=h.className.replace(/\bwf-loading\b/g,"")+" wf-inactive";},config.scriptTimeout),tk=d.createElement("script"),f=false,s=d.getElementsByTagName("script")[0],a;h.className+=" wf-loading";tk.src='//use.typekit.net/'+config.kitId+'.js';tk.async=true;tk.onload=tk.onreadystatechange=function(){a=this.readyState;if(f||a&&a!="complete"&&a!="loaded")return;f=true;clearTimeout(t);try{Typekit.load(config)}catch(e){}};s.parentNode.insertBefore(tk,s)
  })(document);
</script>
<script src="../js/respond.min.js"></script>

</head>
<body>
    <form id="fReceipt" runat="server">
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
             <p><strong>RECEIPT</strong></p>
             <p>for Bob Lindgren's Retirement Dinner<p />
             <p>Friday, June 16, 2017 | 5:00PM<br/>
                Jonathan Club<br/>
                545 S. Figueroa Street<br/>
                Los Angeles, CA 90071</p>

            <div class="divCentered" style="top:140px">
               <p><strong>Reservation for:</strong> <%=customerAddress.firstName %> <%=customerAddress.lastName %> <%=sgName %><br />
                <strong>Meal Selection:</strong> <%=Meal%> <%=sgMeal %><br /><br />
                <strong>***PAYMENT SUMMARY***</strong>
              </p>

              <p><strong>Transaction Date:</strong> <%=DateTime.Now.ToString("MMM dd,yyyy HH:mm")%><br />
                 <strong>Credit Card:</strong> ****<%=CC.Substring(CC.Length - 4)%><br />
                 <strong>Total Amount:</strong> <%=sAmountDue%>
             </p>
      
             <h2><strong>ITEM</strong></h2>
                <div class="ItemText">
                   <% =TransactionMessages%>
              </div>                                 
           </div>
        </div>
    </div>
 </div>
       <br /><br /></section>
 </form>
<div id="subFooter" class="fluid"><p>Bob Lindgren Retirement Dinner | Jonathan Club  | 545 S. Figueroa Street, Los Angeles, CA 90071 | 323.728.9500, Ext. 229</p></div>
<script src="../js/vendor/jquery-1.11.1.min.js"></script>
<script src="../js/vendor/modernizr-2.8.3.min.js"></script>
<script src="../js/plugins.js"></script>
<script src="../js/main.js"></script>

</body>
</html>