<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")
Set rs2 = Server.CreateObject ("ADODB.Recordset")

sql="select * from mas"
rs2.Open sql,conn
%>

<!doctype html>
<html lang="en">
<head>
<meta http-equiv="content-type" content="text/html;  charset=windows-1255">
<meta name="description" content="Manager Waiting List"> 

<link rel="stylesheet" href="desktop.css">
</head>

<body style="text-align:center;">
<img src="/pic2/vlinks.co.png" alt="vlinks logo" title="Go to main page" 
         onclick="window.location='http://www.vlinks.co/vl_manager.asp' " >

<h3>Management System v.1.1</h3>
<h3 style="color:red;font-text:bold">Waiting List:</h3>

<table class="details" style="border: 1px solid black; width:100%; border-collapse:collapse">
  <tr>      
    <th style="background-color: #4CAF50; color: white; border: 1px solid black; text-align:center; padding:15px">Description</th>
    <th style="background-color: #4CAF50; color: white; border: 1px solid black; text-align:center; padding:15px">Web address</th>
    <th style="background-color: #4CAF50; color: white; border: 1px solid black; text-align:center; padding:15px">Delete</th>
  </tr>
<%i=1%>
<%do until rs2.EOF %>

   <tr>
      <td style="border: 1px solid black; text-align:left; padding:15px">
      <td style="border: 1px solid black; text-align:left; padding:15px">
      <td style="border: 1px solid black; text-align:center; width:20px; padding:15px">
      </td>
   </tr>
<%
i=i*-1
rs2.MoveNext 
loop
%>

</table>
<br>
<p>
    <a href="http://www.w3.org/html/logo/">
          <img src="http://www.w3.org/html/logo/badge/html5-badge-h-css3.png" width="133" height="64" alt="HTML5 Powered with CSS3 / Styling" 
                  title="HTML5 Powered with CSS3 / Styling">
     </a> 
</p>
</body>
</html>