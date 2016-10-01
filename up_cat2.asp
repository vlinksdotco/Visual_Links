<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

Set rs=Server.CreateObject ("ADODB.Recordset")

id=Request.QueryString("id")

sql="select * from cat where id="&id
rs.Open sql,conn
%>

<!doctype html>
<html lang="en">
<head>
<title>Update URL</title>
<meta http-equiv="content-type" content="text/html;  charset=windows-1255">
<meta name="viewport" content="user-scalable=no, width=device-width">
<link rel="stylesheet"  type="text/css"
         href="mobile.css" media="only screen and (max-width: 480px)">
<link rel="stylesheet"  type="text/css"
         href="desktop.css" media="screen and (min-width: 481px )">
<!-- [if IE]>
<link rel="stylesheet" type="text/css" href=explorer.css" media="all />
<! [endif] -->
</head>

<body><center>
<img src="pic2/vlinks.co.png" alt="Logo" title="Go to main page" onclick="window.location='vl_manager.asp' " style="cursor:pointer;cursor:hand"><br>
<h3>Management System v. 1.1</h3>
<hr>
 <table border=1>
     <tr>
    <FORM action="up_cat3.asp" method=POST id=form name=form>
     <td>id</td><td><INPUT type="text"  name=id value=<%=rs("id")%> ></td>
    <tr>
    <td>cat name</td><td><INPUT type="text"  name=cat value=<%=rs("cat")%>></td>
    <tr>
     <td>cat name en</td><td><INPUT type="text"  name=la value=<%=rs("lan")%>></td>
    <tr>
     <td>pic</td><td><INPUT type="text"  name=rev value=<%=rs("revnum")%>></td>
    <tr>
    <td colspan=2><INPUT type="submit" value="Submit" id=submit1 name=submit1></td>
 
 </table>
</FORM>
</center>
</body>
</html>

<%
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
%>