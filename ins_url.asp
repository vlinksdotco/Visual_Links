<%@ Language=VBScript %>
<%Response.Buffer =true%>

<%
url=Request.Form("ur") 
rev=Request.Form("rev") 
key=Request.Form("key") 
cat=Request.Form("cat") 
pic=Request.Form("pic")
%>

<br><%=url%>
<br><%=rev%>
<br><%=cat%>

<br><%=lan%>

<%
dim conn1,sql2
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

Set rs = Server.CreateObject ("ADODB.Recordset")
sql="select * from ur where url='"&url&"' and  cat='"&cat&"'"
rs.Open sql,conn

if not rs.EOF then
msg="This url are allreday exsist in tahe same category "
Response.Redirect "new_url.asp?msg="&msg  
end if
	
sql2="insert into ur(url,cat,rev,pic,key) values ('"&url &"','"&cat&"','"&rev&"','"&pic&"','"&key&"')"
conn.Execute sql2

msg="url  insert"
Response.Redirect "new_url.asp?msg="&msg
%>

<%
rs.Close
Set rs=nothing
conn.Close
set conn=nothing
%>