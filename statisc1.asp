<%@ Language=VBScript %>

<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/webtran1.mdb")

tim=request.form("tim")  
cati=request.form("cat")
Set rs = Server.CreateObject ("ADODB.Recordset")
tar=Cint(tim)

sql="SELECT COUNT(ip) as cou FROM cc "
rs.Open sql,conn
cou=rs("cou")
rs.Close 

sql="SELECT SUM(visits) as Total_visits FROM cc WHERE visits>0 "'and  dat < #"&date()+1&"#"
rs.Open sql,conn
Total_visits=rs("Total_visits")
rs.Close 

sql="SELECT SUM(clickc) as Total_click FROM cc WHERE clickc>0"' and  dat < #"&date()+1&"#"
rs.Open sql,conn
Total_click=rs("Total_click")
rs.Close 
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1255"><style>table {    border-collapse: collapse;    border-spacing: 0;    width: 100%;    border: 1px solid #ddd;}th, td {    border: none;    text-align: left;    padding: 8px;}tr:nth-child(even){background-color: #f2f2f2}</style>
</HEAD>
<script>
function ti()
{
  da=select1.value 
  hwnd =window.open('statisc1.asp?tim='+da+',self) ;
}
</script>
<BODY><img src="/pic2/vlinks.co.png" alt="vlinks logo" title="Go to main page"          onclick="window.location='http://www.vlinks.co/vl_manager.asp' " ><h3>Management System v.1.1</h3><div style="overflow-x:auto;">
<table>
<tr>
<td>Total Users</td><td><%=cou%></td>
</tr>
<tr><td>Total Visits</td><td><%=Total_visits%></td></tr>
<tr><td>Total Click</td><td><%=Total_click%></td></tr><tr><td>Average clicks per user</td><td><%=Round(Total_click/Total_visits)%></td></tr>
<FORM action="statisc1.asp" method=post id=form1 name=form1>

<td><SELECT id=tim name=tim>
<OPTION value=7 <%if tim="7" then%>selected<%end if%>>����</OPTION>
<OPTION value=14 <%if tim="14" then%>selected<%end if%>>�������</OPTION>
<OPTION value=30 <%if tim="30" then%>selected<%end if%>>����</OPTION>
<OPTION value=60 <%if tim="60" then%>selected<%end if%>>�������</OPTION>
<OPTION value=90 <%if tim="90" then%>selected<%end if%>>����� ������</OPTION>
<OPTION value=360 <%if tim="360" then%>selected<%end if%>>���</OPTION>
</SELECT></td>
<td>
<SELECT id=cat name=cat>

<OPTION value="ent" <%if cati="ent" then%>selected<%end if%>>������</OPTION>
<OPTION value="click" <%if cati="click" then%>selected<%end if%>>������</OPTION>
<OPTION value="ip" <%if cati="ip" then%>selected<%end if%>>�������</OPTION>

</SELECT>
</td>
<td><INPUT type="submit" value="���" id=submit1 name=submit1></td>
</FORM>
</table></div>
<br>
<%
if tim>0 then
i=1
%><div style="overflow-x:auto;">
<table align=center border=1>
<tr>
<td align=center><%if cati="ent" then%>
    <b><font color=red>���� ������</font>
    </b><%end if%><%if cati="click" then %>
    <b><font color=red>���� ������</font>
    </b><%end if%><%if cati="ip" then%>
    <b><font color=red>���� �������</font>
    </b><%end if%>
</td>
<tr>
<td>
<table  cellpadding=0 cellspacing=1 background="graph.gif" width=600 height=400>
<tr >
<%do until i>tar%>
<%
if(cati="ent")then
sql="SELECT SUM(visits) as Total_visits FROM cc WHERE  dat=#"&date()-(tar-i)&"#"
rs.Open sql,conn
num=rs("Total_visits")
end if

if(cati="click")then

sql="SELECT SUM(clickc) as Total_click FROM cc WHERE  dat=#"&date()-(tar-i)&"#"
rs.Open sql,conn
num=rs("Total_click")
 end if

if(cati="ip")then
sql="SELECT COUNT(ip) as cou FROM cc WHERE  dat=#"&date()-(tar-i)&"#"
rs.Open sql,conn
num=rs("cou")
 end if
%>

<td height=150 valign=bottom   colspan=0 title="<%=num%>,date=<%=date()-(tar-i)%>" style="cursor:pointer;cursor;hand"><div style="height:<%=num%>;width:<%if tar>89 then%>3<%else%>7<%end if%>;background-color:green;opacity:1;filter:alpha(opacity=100)" ></div></td>
<%
rs.close
i=i+1
loop
%>
</table>
</td>
</table></div>
<%end if%>
<%
Set rs=nothing
conn.Close
set conn=nothing
%>
</BODY>
</HTML>
