<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")



Set rs=Server.CreateObject ("ADODB.Recordset")


id=Request.QueryString("id")

sql2="select * from ur where id="&id
rs.Open sql2,conn

temp=rs("click")+1
url=rs("url")
sql = "update ur set click="&temp
sql = sql & " where id="&id

conn2.Execute sql








%>
<%=temp%>
