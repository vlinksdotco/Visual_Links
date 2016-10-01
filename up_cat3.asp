<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

id=Request.Form("id")
cat1=Request.Form("cat") 
lan=Request.Form("la") 
pi=Request.Form("rev") 


sql = "update cat set cat='"&cat1&"',lan='"&lan&"',revnum='"& pi&"' where id="&id 
conn.Execute sql
conn.Close
set conn=nothing
Response.Redirect "vl_manager.asp"
%>
<%=sql%>