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

conn.Execute sql
''''''''''''''''''''''''''''''''''''''''''''

Set conn2 = Server.CreateObject ("ADODB.connection")
conn2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/webtran1.mdb")


tr=Request.ServerVariables("REMOTE_ADDR")

Set rs2 = Server.CreateObject ("ADODB.Recordset")
sql3="select * from cc where ip='"&tr&"' and dat=#"&Date()&"#"
rs2.Open  sql3,conn2
 
  cou=rs2("clickc")+1
  sql="UPDATE cc set clickc="&cou&" where ip='"&tr&"' and dat=#"&Date()&"#"
  conn2.Execute sql
 




''''''''''''''''''''''''''''''''''''''''''''''
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
rs2.Close
set rs2=nothing
conn2.Close 
set conn2=nothing



%>

