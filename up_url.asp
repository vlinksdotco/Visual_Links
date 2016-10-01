<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

id=Request.Form("id")
url=Request.Form("ur") 
rev=Request.Form("rev") 
key=Request.Form("key") 
cat=Request.Form("cat") 
pic=Request.Form("pic")
click=Request.Form("cli")

sql = "update ur set url='"&url&"',rev='"&rev&"' ,cat='"&cat&"', key='"&key&"', pic='"&pic&"',click='"&click&"' where id="&id 

conn.Execute sql
conn.Close 
set conn=nothing
Response.Redirect "up_ur.asp"%>