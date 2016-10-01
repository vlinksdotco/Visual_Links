<%@ Language=VBScript %>
<%
dim conn1,sql2,temp
count=0
count2=0
temp=0
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

Set rs = Server.CreateObject ("ADODB.Recordset")
sql="select * from cat "
rs.Open sql,conn

cat=Request.QueryString("cat")
if cat="" then
temp=1
else

Set rs2 = Server.CreateObject ("ADODB.Recordset")
sql="select * from ur where cat='"&cat&"'"
rs2.Open sql,conn

end if
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
<img src="/pic2/vlinks.co.png" alt="Logo" title="Go to main page"
          onclick="window.location='vl_manager.asp' "style="cursor:pointer;cursor:hand"><br>
<h3>Management System v. 1.0.2</h3>
<h4>Select the category & icon then  update the data</h4>
<table style="text-align:center; /*border-style:outset;*/">
  <tr height=10>
		<table  style="text-align:center; border-style:outset"> <!-- dir=rtl> -->
		<tr height=10 align=center>

		<%do until rs.EOF %>
		<td width=100 align=center style="border-style:0;
		                                  background-color:#00008b;
		                                  border-width:1pix;
		                                  color:#d3d3d3;cursor:pointer;
		                                  cursor:hand;"   onmouseover='this.style.backgroundColor="#4169e1";this.style.color="#ffd700"' onmouseout='this.style.backgroundColor="#00008b";this.style.color="#d3d3d3"'
		onclick="window.location='up_ur.asp?cat=<%=rs("cat")%> '"><%=rs("cat")%>
		</td>

<%
count=count+1
if count=8 then
%>
		<tr height=10 align=center>
<%
count=0
end if
%>

<%
rs.MoveNext 
loop
%>
		</table>
<%if temp=0 then%>
		<tr>
	<td>
<table  align=center  dir=rtl>
	<tr height=10 align=center >

<%do until rs2.EOF %>
	<td>
		<img src=pic2/<%=rs2("pic")%> title="<%=rs2("rev")%>" 
		     onclick="window.location='up_ur2.asp?id=<%=rs2("id")%>'" border="1"  			  style="cursor:pointer;cursor:hand">
	</td>



<%
count2=count2+1
if count2=5 then
%>
	<tr height=10 align=center>

<%
count2=0
end if
%>

<%
rs2.MoveNext 
loop
%>
	</table>
	</td>
<%end if%>
</table>
</body>
</html>

<%
rs.Close
Set rs=nothing
conn.Close
set conn=nothing
%>
