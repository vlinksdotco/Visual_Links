<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")
Set rs2 = Server.CreateObject ("ADODB.Recordset")

SearchText=request.Form("text1")

Set rs2 = Server.CreateObject ("ADODB.Recordset")

sql2="SELECT * FROM ur  WHERE rev like '%"&SearchText&"%' or key like '%"&SearchText&"%' or url like '%"&SearchText&"%' or cat like '%"&SearchText&"%'"

rs2.Open sql2,conn
%>

<!doctype html>
<html lang="en">

<head>
<title>Update URL - search</title>
<meta http-equiv="content-type" content="text/html;  charset=utf-8">
<meta name="viewport" content="user-scalable=yes, width=device-width">
<link rel="stylesheet"  type="text/css"
         href="mobile.css" media="only screen and (max-width: 480px)">
<link rel="stylesheet"  type="text/css"
         href="desktop.css" media="screen and (min-width: 481px )">
<!-- [if IE]>
<link rel="stylesheet" type="text/css" href=explorer.css" media="all />
<! [endif] -->
 </head>
 
<body>
<center>
<img src="/pic2/vlinks.co.png" alt="Logo" 
          title="Go to main page"  onclick="window.location='vl_manager.asp' " >
<br>
<hr>
<h3>Management System v. 1.1</h3>
<hr>
<form name="searchBox" method="post">
 <table>
  <tr>	   
   <td>
	<input type="text"  name="text1"  <%if SearchText="" then %>value=""  <%else%>value=" <%=SearchText%>"<%end if%> 
                   style="height:30;width:70;font-size:20" 
                   placeholder="Enter here what you want to find" autofocus >
   </td>
   <td>
	<input type="submit" value="Search">
   </td>
  </tr>		  
 </table>
</form>

<table>
  <tr>
   <%if not SearchText="" then%>
   <% 
k=0
do until rs2.EOF 
%>

   <td>
      <img src=pic2/<%=rs2("pic")%> title="<%=rs2("rev")%>" onclick="window.location='up_ur2.asp?id=<%=rs2("id")%>'" onmouseover="window.status='<%=rs2("url")%>'" onmouseout="window.status=''">
   </td> 

<%if k=4 then%><tr>
<%end if%>
<%
rs2.MoveNext 

k=k+1
if k>4 then k=0 end if
loop
%>	
<%end if%>
   
  </table>
</body>
</html>

<%
rs2.Close
Set rs2=nothing
conn.Close
set conn=nothing
%>