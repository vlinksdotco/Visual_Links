<%@ Language=VBScript %>

<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")
Set rs2 = Server.CreateObject ("ADODB.Recordset")

sql="select count(id) as mass from  mas"
rs2.Open sql,conn
mas_qty=rs2("mass")
rs2.Close

sql="select count(id) as ur from ur "
rs2.Open sql,conn
web_qty=rs2("ur")
rs2.Close
Set rs2=nothing
conn.Close
set conn=nothing
%>

<!doctype html>
<html lang="en">
<head>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.3/jquery.min.js"></script>

<script>
alert("\nWelcome to VLinks! V is for Visual !\nThe best Visual, Social and Anonymous search engine.\n");

// jquery code ...
$(document).ready(function(){
    $("button").click(function(){ //hide div#ads or hide all page with ("*").hide
        $("#ads").toggle(); // do on/off with toggle
    });
});
$(document).ready(function(){    $("img").click(function(){        $(this).hide();    });});
function about(){
  alert('VL Manager v. 1.0.3\nWritten by Visual Links team.');
}
/*
var today = new Date().getDay(); // print number of day
document.write(today)

var answer = confirm('\nWelcome to vlinks.co ! V is for Visual !\nThe best visual search engine.\n\nAre you agree with it?');

if(answer){
	var name = prompt('\nWhat is your name?');
	if(name){
		alert('\nHello ' + name + ' !' + '\nJoin us !');
	}
	else{
		alert('\nYou can always stay anonymous :-) !');
	}
}
else{
	alert('\nWe are still evolving ! Give it a chance !');
}
*/
</script>

<title>VL Manager</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<meta name="viewport" content="user-scalable=no, width=device-width">
<meta name="robots" content="noindex">

<link rel="stylesheet" type="text/css" href="mobile.css" media="only screen and (max-width: 480px)">
<link rel="stylesheet" type="text/css" href="desktop.css" media="screen and (min-width: 481px)">
</head>

<body style="text-align:center;">
<header> 
<div style="float:left;">
<img src="pic2/vlinks.co.png"   alt="Visual Links's Logo" title="Show me main page" 
    style="position:absolute; border:1; border-color:blue; box-shadow:3px 3px 2px #585858;"     
    onclick="window.location='index.asp'">
</div>
</header>
<h1>Management System v. 1.0.4</h1>

<%
'date
Function myfunction()
 myfunction=Date()
End Function

response.write("Date : " & myfunction() & "&nbsp")
'print what day
d=weekday(Date)
%>
<div id="adv/info">
<button>Ads ON/OFF</button>
<button onclick="about()">About</button>
</div>
<!-- ************************* Advertisers rubric starts here ************************* --><h5>OUR ADVERTISERS</h5>
<div id="ads" style="height:100px; width:1260px; border:1px solid black; background-color:white; margin-top:10px;">
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="a1">
      <img src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow:3px 3px 2px #585858; margin-top:7px;"
        onclick='aclick("openpage.asp?id=2625")' id="a2">
  </a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="b1">
      <img src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858;"
        onclick='aclick("openpage.asp?id=2625")' id="b2">
   </a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="c1">
      <img src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858;"
        onclick='aclick("openpage.asp?id=2625")' id="c2">
   </a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="d1">
      <img src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858;"
        onclick='aclick("openpage.asp?id=2625")' id="d2">
   </a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="e1">
      <img src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858;"
        onclick='aclick("openpage.asp?id=2625")' id="e2">
   </a>
</div>
<h2>Current status :</h2>
<h2 style="color:green;"><%=web_qty%>&nbsp; websites &nbsp; / &nbsp; <%=mas_qty%>&nbsp; messages</h2>
<div id="buttons" style="height:100px; width:1260px;">  
<input type="button" value="Waiting list" onclick="window.location='open_mass.asp'" id=button1 name=button1>
<input type="button" value="Todo list"  onclick="window.location='todo.html'" id=button4 name=button4>
<input type="button" value="Search" onclick="window.location='search.asp'" id=button2 name=button2>
<!--img src="images/search_button2.png"   alt="Search button" onclick="window.location='search.asp'"-->
<input type="button" value="Add" onclick="window.location='new_url.asp'" >
<input type="button" value="Update"  onclick="window.location='up_ur.asp'">
<input type="button" value="Update category"  onclick="window.location='up_cat.asp'" id=button3 name=button3>
<input type="button" value="Statistics" onclick="window.location='statisc.asp'" id=button5 name=button5>
<!--input type="button" value="Counter" onclick="window.location='statisc1.asp'" id=button6 name=button6-->  
</div>

<div id="validator">
<a href="http://validator.w3.org/check?uri=http%3A%2F%2Fwww.vlinks.co%2Fmanager.asp">
<img src="http://www.w3.org/html/logo/badge/html5-badge-v-solo.png" alt="HTML5 Powered" title="HTML5 Powered">
</a>
</div>
</body>
</html>