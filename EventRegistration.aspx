<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="EventRegistration.aspx.cs" Inherits="LouACH.EventRegistration" %>

<!doctype html>
<!--[if lt IE 7]> <html class="ie6 oldie"> <![endif]-->
<!--[if IE 7]>    <html class="ie7 oldie"> <![endif]-->
<!--[if IE 8]>    <html class="ie8 oldie"> <![endif]-->
<!--[if gt IE 8]><!-->
<html class="">
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
    <script>
        (function (d) {
            var config = {
                kitId: 'jfz5dcg',
                scriptTimeout: 3000,
                async: true
            },
            h = d.documentElement, t = setTimeout(function () { h.className = h.className.replace(/\bwf-loading\b/g, "") + " wf-inactive"; }, config.scriptTimeout), tk = d.createElement("script"), f = false, s = d.getElementsByTagName("script")[0], a; h.className += " wf-loading"; tk.src = 'https://use.typekit.net/' + config.kitId + '.js'; tk.async = true; tk.onload = tk.onreadystatechange = function () { a = this.readyState; if (f || a && a != "complete" && a != "loaded") return; f = true; clearTimeout(t); try { Typekit.load(config) } catch (e) { } }; s.parentNode.insertBefore(tk, s)
        })(document);
    </script>
    <script src="../js/respond.min.js"></script>
    <script>
     function ShowHideDiv(CheckGuest) {
         var guest = document.getElementById("guest");
         var button = document.getElementById("ButtonReserve");
         var guestname = document.getElementById("txtGuestName");
         var offRadio = document.getElementById("Radiog4");
         guest.style.display = CheckGuest.checked ? "block" : "none";
         button.style.display = CheckGuest.checked ? "none" : "block";
         guestname.value = '';
         offRadio.checked = true;
     }
     function ShowHideButton() {
         var button = document.getElementById("ButtonReserve"); 
         button.style.display = "block";
     }
     function ShowHideGuest(RadioB) {
         var button = document.getElementById("ButtonReserve");
         button.style.display = "block";
     }
     function GoToNext() 
     {
         var FName,LName ,Company, Email, emailExp, cGuest,gName;

         FName = document.getElementById('txtFName').value;
         LName = document.getElementById('txtLName').value;
         Company = document.getElementById('txtCoName').value;
         Email = document.getElementById('txtEMail').value;
         emailExp = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
         cGuest = document.getElementById('CheckGuest').checked;
         gName = document.getElementById('txtGuestName').value;
         if (FName == '' && LName == '' && Company == '' && Email == '') 
         {
             alert('Enter All Fields');
             return false;
         }
         if (FName == '') 
         {
             alert('Please Enter First Name');
             return false;
         }
         if (LName == '') 
         {
             alert('Please Enter Last Name');
             return false;
         }
         if (Company == '') 
         {
             alert('Please Enter Company or NA');
             return false;
         }

         if (Email == '')
         {
             alert('Email Is Required');
             return false;
         }
         if (Email != '')
         {
             if (!Email.match(emailExp))
             {
                 alert('Invalid Email');
                 return false;
             }
         }
         if (cGuest  && gName=='') {
             alert('Please Enter Guest Name');
             return false;
         }
         fRegister.action = 'RegistrationPay.aspx';
         fRegister.submit();
         }

 
</script>
</head>
<body>
    <form id="fRegister" runat="server">
        <header class="fluid orange">
            <div class="gridContainer clearfix">
                <div id="div1" class="fluid">
                    <img src="images/RetirementDinnerHeader.png" alt="banner"><br>
                    <br>
                </div>
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
                        <div class="fluid">
                            <p><strong>R.S.V.P.</strong></p>
                            <div>
                                <p>
                                    <label for="txtCoName" class="HGLabel">Your Name:</label>
                                    <input class="TxtArea" id="txtFName" name="txtFName" style="width: 220px" placeholder="First Name" />&nbsp;<input class="TxtArea" id="txtLName" name="txtLName" style="width: 220px" placeholder="Last Name" /><br />
                                    <label for="txtCoName" class="HGLabel">&nbsp;&nbsp;&nbsp;Company:</label>
                                    <input class="TxtArea" name="txtCoName" id="txtCoName" style="width: 445px" /><br />
                                    <label for="txtCoName" class="HGLabel">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Email:</label>
                                    <input class="TxtArea" id="txtEMail" name="txtEMail" style="width: 445px" />
                                </p>
                                <p>
                                    <label class="RadioLabel"><strong>Select Meal:</strong></label><br>
                                    <input id="Radio1" type="radio" name="rMeal" value="Filet Mignon" onchange="ShowHideButton()" />
                                    <label class="RadioLabel">Filet Mignon  </label>
                                    &nbsp;&nbsp;<input id="Radio2" type="radio" name="rMeal" value="Chilean Sea Bass" onchange="ShowHideButton()" />
                                    <label>Chilean Sea Bass</label>
                                    &nbsp;&nbsp;<input id="Radio3" type="radio" name="rMeal" value="Vegetarian Pasta" onchange="ShowHideButton()" />&nbsp;<label class="RadioLabel">Vegetarian Pasta</label>
                                
                                </p>
                                <p>
                                    <input id="CheckGuest" name="CheckGuest" type="checkbox" value="1" onchange="ShowHideDiv(this)" />
                                    <label class="QuestionLabel">I will be bringing spouse/guest ($200 charge). </label>
                                </p>

                                <div id="guest" style="display: none">
                                    <p>
                                        <label for="txtCoName">Guest Name:</label>
                                        <input class="TxtArea" id="txtGuestName" name="txtGuestName" value="" style="width: 435px" /></p>
                                    <p>
                                        <label class="RadioLabel"><strong>Select Meal:</strong></label><br />
                                        <input id="Radiog1" type="radio" name="gMeal" value="Filet Mignon" onchange="ShowHideButton()" />
                                        <label class="RadioLabel">Filet Mignon  </label>
                                        &nbsp;&nbsp;<input id="Radiog2" type="radio" name="gMeal" value="Chilean Sea Bass" onchange="ShowHideButton()" />&nbsp;<label class="RadioLabel">Chilean Sea Bass</label>&nbsp;&nbsp;<input id="Radiog3" type="radio" name="gMeal" value="Vegetarian Pasta" onchange="ShowHideButton()" />
                                        <label>Vegetarian Pasta</label><br />
                                             <input id="Radiog4" type="radio" name="gMeal" value="" style="visibility:hidden"/>
                                    </p>
                                </div>

                                <div>
                                    <input id="ButtonReserve" type="button" class="button" value="R.S.V.P." style="display:none" onclick="GoToNext()" />
                                </div>
                            </div>

    <p><span style="color: #ed1c24; font-size: 10px;">Cancellation Policy: PIASC must receive cancellations 48 hours prior to the event to receive a full refund.</span></p>
    </div>
    </div>
    </div>
    </div><br />
    <br />
    </section>
    <div id="subFooter" class="fluid">
        <p>Bob Lindgren Retirement Dinner | Jonathan Club  | 545 S. Figueroa Street, Los Angeles, CA 90071 | 323.728.9500, Ext. 229</p>
    </div>
    <script src="../js/vendor/jquery-1.11.1.min.js"></script>
    <script src="../js/vendor/modernizr-2.8.3.min.js"></script>
    <script src="../js/plugins.js"></script>
    <script src="../js/main.js"></script>
        </form>
</body>
</html>