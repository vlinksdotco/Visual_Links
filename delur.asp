<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")
id=Request.QueryString("id")

sql = "delete * from ur  where id="&id
 conn.Execute sql

conn.Close 
set conn=nothing
Response.Redirect "vl_manager.asp"

%>
