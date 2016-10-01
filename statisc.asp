<%@ Language=VBScript%><%Response.Buffer = True%>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/webtran1.mdb")

Set rs = Server.CreateObject ("ADODB.Recordset")

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


sql="select * from cc"
rs.Open sql,conn
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0"><meta http-equiv="content-type" content="text/html;  charset=windows-1255"><title>Statistics</title><link rel="stylesheet" href="desktop.css"><style>* {  box-sizing: border-box;}#myInput {  background-image: /* url('/css/searchicon.png');*/  background-position: 10px 10px;  background-repeat: no-repeat;  width: 100%;  font-size: 16px;  padding: 12px 20px 12px 40px;  border: 1px solid #ddd;  margin-bottom: 12px;}#myTable {  border-collapse: collapse;  width: 100%;  border: 1px solid #ddd;  font-size: 18px;}#myTable th, #myTable td {  text-align: left;  padding: 12px;}#myTable tr {  border-bottom: 1px solid #ddd;}#myTable tr.header, #myTable tr:hover {  background-color: #f1f1f1;}table {    border-collapse: collapse;    border-spacing: 0;    width: 100%;    border: 1px solid #ddd;}th, td {    border: none;    text-align: left;    padding: 8px;}tr:nth-child(even){background-color: #f2f2f2}</style>
</HEAD>
<BODY>
<center><body style="text-align:center;"><img src="/pic2/vlinks.co.png" alt="vlinks logo" title="Go to main page"          onclick="window.location='http://www.vlinks.co/vl_manager.asp' " ><h3>Management System v.1.1</h3>
<a href="http://m.maploco.com/details/e4048v6c">
  <img style="border:0px;" src="http://www.maploco.com/vmap/7132116.png" 
           alt="Locations of Site Visitors" title="Locations of Site Visitors"/>
</a><h3><%=date()%></h3>
<div style="overflow-x:auto;"><input type="text" id="myInput" onkeyup="myFunction()" placeholder="Search for dates..." title="Type in a date">
<table id="myTable"><tr class="header">
<th style="background-color: #4CAF50; color: white; border: 1px solid black; text-align:center; padding:15px">IP address</th><th style="background-color: #4CAF50; color: white; border: 1px solid black; text-align:center; padding:15px">Enters</th><th style="background-color: #4CAF50; color: white; border: 1px solid black; text-align:center; padding:15px">Clicks</th><th style="background-color: #4CAF50; color: white; border: 1px solid black; text-align:center; padding:15px">Date</th></tr>
<%da=rs("dat")%>
<%do until rs.EOF %>
<%response.flush
if not da=rs("dat")then
%>
<tr>
<td colspan=4 height=20  bgcolor=LimeGreen></td>

<%
da=rs("dat")

end if%>
<tr>
<td style="text-align:center"><%=rs("ip")%></td>
<td style="text-align:center"><%=rs("visits")%></td>
<td style="text-align:center"><%=rs("clickc")%></td>
<td style="text-align:center"><%=rs("dat")%></td>


<%rs.MoveNext 
loop
%>
</table>
</table></div><div style="overflow-x:auto;">
<table>
<tr>
<td>Total Users</td><td><%=cou%></td>
<tr>
<td>Total Visits</td><td><%=Total_visits%></td>
<tr>
<td>Total Click</td><td><%=Total_click%></td>
<tr>
<td>Average clicks per user</td><td><%=Total_click/Total_visits%></td>

</table>
</div><div style="overflow-x:auto;">
<table>

<%
rs.Close

k=0
i=-4
do while k<5

sql="SELECT SUM(visits) as Total_visits FROM cc WHERE visits>0 and  dat=#"&date()+i&"#"
rs.Open sql,conn
Total_visits=rs("Total_visits")

if rs.EOF  then Total_visits=0 end if
rs.Close 

sql="SELECT SUM(clickc) as Total_click FROM cc WHERE clickc>0and  dat=#"&date()+i&"#"
rs.Open sql,conn
Total_click=rs("Total_click")
rs.Close 

%>
<tr>
<td><%=date()+i%></td><td  width=150><div style="height:20;width:<%=Total_visits%>;background-color:red"></div></td ><td width=40 align=center><%=Total_visits%></td>
<tr>
<td></td><td width=150><div style="height:10;width:<%=Total_click%>;background-color:green"></div></td><td width=40 align=center><%=Total_click%></td>

<%
k=k+1
i=i+1
loop%>
</table></div>
<%
Set rs=nothing
conn.Close
set conn=nothing
%>
</center><script>function myFunction() {  var input, filter, table, tr, td, i;  input = document.getElementById("myInput");  filter = input.value.toUpperCase();  table = document.getElementById("myTable");  tr = table.getElementsByTagName("tr");  for (i = 0; i < tr.length; i++) {    td = tr[i].getElementsByTagName("td")[0];//    if (td) {      if (td.innerHTML.toUpperCase().indexOf(filter) > -1) {        tr[i].style.display = "";      } else {        tr[i].style.display = "none";      }    }  }}</script>
</BODY>
</HTML>