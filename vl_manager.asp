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

function about(){
  alert('VL Manager v. 1.0.1\nWritten by Visual Links team.');
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
<meta http-equiv="content-type" content="text/html;  charset=utf-8">
<meta name="viewport" content="user-scalable=no, width=device-width">
<meta name="robots" content="noindex">

<link rel="stylesheet" type="text/css" href="mobile.css" media="only screen and (max-width: 480px)">
<!--link rel="stylesheet" type="text/css" href="desktop.css" media="screen and (min-width: 481px )"-->
</head>

<body style="text-align:center;">
<header> 
<div style="float:left;">
<img src="pic2/vlinks.co.png"   alt="Visual Links's Logo" title="Show me main page" 
    style="position:absolute; border-color:blue; box-shadow:3px 3px 2px #585858;"     
    onclick="window.location='index.asp'">
</div>
</header>
<h1>Management System v. 1.2</h1>
<%
'date
Function myfunction()
 myfunction=Date()
End Function

response.write("Date : " & myfunction() & "&nbsp")

i=hour(time)
If i < 10 Then
   response.write("Good morning!" & "&nbsp")
Else
   response.write("Have a nice day!" & "&nbsp")
End If

'print what day
d=weekday(Date)

Select Case d
  Case 1
    response.write("Sleepy Sunday" & "<br><br>")
  Case 2
    response.write("Monday again!" & "<br><br>")
  Case 3
    response.write("Just Tuesday!" & "<br><br>")
  Case 4
    response.write("Wednesday!" & "<br><br>")
  Case 5
    response.write("Thursday..." & "<br><br>")
  Case 6
    response.write("Finally Friday!" & "<br><br>")
  Case Else
    response.write("Super Saturday!!!!" & "<br><br>")
End Select
%>
<p>
<button>Ads ON/OFF</button>
</p>
<p>
<button onclick="about()">About</button>
</p>
<!-- ************************* Advertisers rubric starts here ************************* -->
<div id="ads">
<h5>OUR ADVERTISERS</h5>
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="a1">
      <img src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858;"
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
<h1>Current status :</h1>
<h2 style="color:green;"><%=web_qty%>&nbsp; websites</h2> 
<h2 style="color:red;"><b><%=mas_qty%></b>&nbsp; messages</h2>
<div id="buttons">  
<input type="button" value="Waiting list" onclick="window.location='open_mass.asp'" id=button1 name=button1>
<p>
<input type="button" value="Todo list"  onclick="window.location='todo.html'" id=button4 name=button4>
<input type="button" value="Search" onclick="window.location='search.asp'" id=button2 name=button2>
<!--img src="images/search_button2.png"   alt="Search button" onclick="window.location='search.asp'"-->
<input type="button" value="Add" onclick="window.location='new_url.asp'" >
<input type="button" value="Update"  onclick="window.location='up_ur.asp'">
</p>
<!--input type="button" value="Update category"  onclick="window.location='up_cat.asp'" id=button3 name=button3-->
<input type="button" value="Statistics" onclick="window.location='statisc.asp'" id=button5 name=button5>
<input type="button" value="Counter" onclick="window.location='statisc1.asp'" id=button6 name=button6>  
</div>
<div id="validator">
<a href="http://validator.w3.org/check?uri=http%3A%2F%2Fwww.vlinks.co%2Fmanager.asp">
<img src="http://www.w3.org/html/logo/badge/html5-badge-v-solo.png" alt="HTML5 Powered" title="HTML5 Powered">
</a>
</div><audio controls autoplay loop>
  <source src="/sounds/HappyPills.mp3" type="audio/mpeg">
</audio>
</body>
</html>