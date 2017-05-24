<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EventMakePayment.aspx.cs" Inherits="LouACH.EventMakePayment" %>
<<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
  <meta content="text/html; charset=ISO-8859-1"  http-equiv="content-type"/>
<title></title>
<script src="../js/respond.min.js"></script>
<script type="text/javascript">
    function GoToNext() {
        var Address, City, State, Zip, CCNo, CCExp, CCCode;

        Address = document.getElementById('txtAddress').value;

        City = document.getElementById('txtCity').value;
        State = document.getElementById('txtState').value;
        Zip = document.getElementById('txtZip').value;
        CCNo = document.getElementById('txtCC').value;
        CCExp = document.getElementById('txtExpDt').value;
        CCCode = document.getElementById('txtSecNo').value;

        if (Address == '' && City == '' && State == '' && Zip == '' && CCNo == '' && CCExp == '' && CCCode == '') {
            alert('Enter All Fields');
            return false;
        }
        if (Address == '') {
            alert('Please Enter Address');
            return false;
        }
        if (City == '') {
            alert('Please Enter City');
            return false;
        }
        if (State.length != 2) {
            alert('Please Enter State Abbreviation');
            return false;
        }

        if (Zip == '' || isNaN(Zip)) {
            alert('Please enter Zip Code');
            return false;
        }

        if (CCNo == '' || isNaN(CCNo)) {
            alert('Please Credit Card No (no spaces)');
            return false;
        }
        //if (CCExp == '' || CCExp.substring(2, 3) !== '/') {
        if (CCExp == '' || isNaN(CCExp)) {
            alert('Please Credit Card Expiration date (mmyy)');
            return false;
        }
        if (CCCode == '' || isNaN(CCCode)) {
            alert('Please Credit Card security code');
            return false;
        }

        fPayment.action = 'EventReceipt.aspx';
        fPayment.submit();
        var spinner = document.getElementById("spinner");
        spinner.style.display = "block";
    }
</script>
 
</head>
<body>
    <form id="fPayment" runat="server">
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
<p><strong>PAYMENT PAGE</strong></p>

        <div class="box2">
        <p><label for="txtAmount">Amount Due:</label><input class="TxtArea" name="txtAmount" id="txtAmount" ="10" value="<%=sAmountDue %>"/><br/><br/>
        <label for="txtCoName">&nbsp;&nbsp;&nbsp;&nbsp;Name:</label><input class="TxtArea" name="txtFName" value="<%=fName %>" style="width:220px" placeholder="First Name"/><input class="TxtArea" name="txtLName" value="<%=lName %>" style="width:220px"  placeholder="Last Name"/><br/>
        <label for="txtAddress">Address:</label><input class="TxtArea" name="txtAddress"  id="txtAddress" value="" style="width:450px"/><br />
        <label for="txtCity">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;City:</label><input class="TxtArea" name="txtCity"  id="txtCity" value=""/ style="width:450px"/><br />
        <label for="txtState">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;State:</label><input class="TxtArea" name="txtState"  id="txtState" value="" style="width:50px"/>
        <label for="txtZip">Zip:</label><input class="TxtArea" name="txtZip"  id="txtZip" value="" style="width:100px"/><br /><br />
        <label for="txtCC">Card Number:</label><input class="TxtArea" name="txtCC"  id="txtCC" value="4111111111111111"  style="width:250px"/><br />
        <label for="txtExpDt">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Expiration:</label><input class="TxtArea" name="txtExpDt"  id="txtExpDt" value=""  style="width:75px"  placeholder="  mmyy"/><label for="txtSecNo"> Security Code:</label><input class="TxtArea" name="txtSecNo"  id="txtSecNo" value="" style="width:75px"/><br />
            <input id="ButtonPay" type="button" value="Make Payment" class="button" onclick="GoToNext()" /><div class="loader" id="spinner"  style="display: none"></div>
             
          </div>  
          

     </div></div><br /><br /></section>
        </form>
</body>

</html>
