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
<meta name="viewport" content="user-scalable=no, width=device-width"><link rel="stylesheet"  type="text/css" href="desktop.css" media="screen and (min-width:481px)">
<link rel="stylesheet"  type="text/css" href="mobile.css" media="only screen and (max-width:480px)">
</head>

<body style="text-align:center;">
<br>
<img src="/pic2/vlinks.co.png" alt="Logo" title="Go to main page"
          onclick="window.location='vl_manager.asp'" style="border: 1px solid black; cursor:pointer;cursor:hand"><br>
<h3>Management System v. 1.0.2</h3>
<h4>Select the category & icon then  update the data</h4>
<table class="details" style="border: 1px solid black; width:100%; border-collapse:collapse">
  <tr>
		<table class="details" style="border: 1px solid black; width:100%; border-collapse:collapse">
		<tr>

		<%do until rs.EOF %>
		<td style="border-style:0; background-color:blue; border-width:1px; color:yellow; cursor:pointer; border:1; cursor:hand;" onmouseover='this.style.backgroundColor="red";this.style.color="black"' 
                onmouseout='this.style.backgroundColor="blue";this.style.color="yellow"'
		     onclick="window.location='up_ur.asp?cat=<%=rs("cat")%> '"><%=rs("lan")%>
		</td>

<%
count=count+1
if count=8 then
%>
		<tr>
<%
count=0
end if
%>

<%
rs.MoveNext 
loop
%>
		</table> <br>
<%if temp=0 then%>

<%do until rs2.EOF %>
<% j=0 %>
		<img src=pic2/<%=rs2("pic")%> title="<%=rs2("rev")%>" 
		     onclick="window.location='up_ur2.asp?id=<%=rs2("id")%>'"	>	                  <!--style="border:1px solid blue; box-shadow: 3px 3px  2px #585858; border: 1; cursor:pointer;  cursor:hand"-->
<%if j=4 then%><br><br><%end if%>

<%
count2=count2+1
if count2=5 then
%>

<%
count2=0
end if
%>

<%
rs2.MoveNext 
loop
%>
	
<%end if%>

</body>
</html>

<%
rs.Close
Set rs=nothing
conn.Close
set conn=nothing
%>