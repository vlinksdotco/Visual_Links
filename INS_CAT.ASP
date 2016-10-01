<%@ Language=VBScript %>
<%Response.Buffer =true%>

<%
logNA=Request.Form("cat") 
%>

<%
dim conn1,sql2
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

Set rs = Server.CreateObject ("ADODB.Recordset")
sql="select * from cat where cat='"&logNA&"'"
rs.Open sql,conn

if not rs.EOF then
msg="category exist"
Response.Redirect "new_cat.asp?msg="&msg  
end if
	
sql2="insert into cat(cat)values ('"&logNA&"')"
conn.Execute sql2

msg="category insert"
 
Response.Redirect "new_cat.asp?msg="&msg
%>
<%
rs.Close
Set rs=nothing
conn.Close
set conn=nothing
%>