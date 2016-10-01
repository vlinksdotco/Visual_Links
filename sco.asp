<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/gam.mdb")

id=Request.QueryString("id") 
sco=Request.QueryString("sco")

Set rs2 = Server.CreateObject ("ADODB.Recordset")

sql2="select * from gam where id="&id
rs2.Open sql2,conn

player=rs2("dirog")+1
scor=rs2("score")+sco

sql="UPDATE gam set score="&scor&",dirog="&player&" where id="&id
conn.Execute sql

rs2.Close
Set rs2=nothing
conn.Close
set conn=nothing
%>