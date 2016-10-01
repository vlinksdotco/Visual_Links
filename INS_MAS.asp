<%@ Language=VBScript %>
<%
mas=Request.Form("mas") 
ur=Request.Form("ur") 
my="38"
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")
if not ur="" then
  sql="insert into mas(url,mem)values('"&ur&"','"&mas&"')"
else
  sql="insert into mas(mem)values('"&mas&"')"
end if

conn.Execute sql         
conn.Close
set conn=nothing
msg="Thanks for the new link !" 
Response.Redirect "index.asp?msg="&msg
%>
