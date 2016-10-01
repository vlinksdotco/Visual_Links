<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

Set rs=Server.CreateObject ("ADODB.Recordset")

id=Request.QueryString("id")

sql2="select * from ur where id="&id
rs.Open sql2,conn
%>

<!doctype html>
<html lang="en">
<head>
<meta http-equiv="content-type" content="text/html; charset=windows-1255">
<title>Management System</title>
</head>
<body>
<center>
<img src="pic2/vlinks.co.png" alt="Logo" title="Go to main page" onclick="window.location='vl_manager.asp' " style="cursor:pointer;cursor:hand"><br>
<h3>Management System v. 1.0.2</h3>
<hr>
<h4>Update the icon's data</h4> 
<table border=1>
     <tr>
    <FORM action="up_url.asp" method=POST id=form name=form>
     <td>id</td><td><input type="text"  name=id value=<%=rs("id")%>  size=30></td>
    <tr>
    <td>Web site address</td><td><input type="text"  name=ur value="<%=rs("url")%>" size=30></td>
    <tr><td>Keyword </td><td><textarea rows=6 cols=30 name=key ><%=rs("key")%></textarea>
    <tr><td>Description</td><td><textarea rows=6 cols=30 name=rev ><%=rs("rev")%></textarea></td>
    <tr><td>Icon's filename</td><td><input type="text"  name=pic value="<%=rs("pic")%>" size=30></td>
    <tr><td>Category</td><td><input type="text"  name=cat value="<%=rs("cat")%>" size=30></td>
    <tr><td>Clicks</td><td><input type="text"  name=cli value=<%=rs("click")%> size=30></td>
    <tr><td colspan=2><input type="submit" value="Submit" id=submit1 name=submit1></td>
    <tr><td colspan=2><input type="button" value="Delete" onclick="window.location='delur.asp?id=<%=rs("id")%>'"></td>
</table>
</FORM>
</body>
</html>

<%
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
%>