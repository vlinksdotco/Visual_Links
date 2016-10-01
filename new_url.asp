<%@ Language=VBScript %>
<%
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

Set rs2 = Server.CreateObject ("ADODB.Recordset")

sql2="select * from cat"
rs2.Open sql2,conn
%>

<!doctype html>
<html lang="en">
<head>
<title>Add URL - new_url
</title>

<meta http-equiv="content-type" content="text/html;  charset=windows-1255">
<meta name="viewport" content="user-scalable=no, width=device-width">

<link rel="stylesheet"  type="text/css"
         href="mobile.css" media="only screen and (max-width: 480px)">
<link rel="stylesheet"  type="text/css"
         href="desktop.css" media="screen and (min-width: 481px )">
<!-- [if IE]>
<link rel="stylesheet" type="text/css" href=explorer.css" media="all />
<! [endif] -->

<script type="text/javascript">
var t=1;

function Valid2(str){
  r=""
  var newstr="";
  var a,count;
  a=str.split("'");
  count=a.length;
  for(i=0; i<count; i++)
    newstr=newstr+a[i];
		
  a=newstr.split(",");
  count=a.length;
  for(i=0; i<count; i++)
    r=r+a[i];		
  return r ;
}
function log(){
  document.form.ur.value=Valid2(document.form.ur.value)l

  if(document.form.ur.value==""){
    alert("Enter url");
    t=0;
   }   
  return;
}
function rev(){
  document.form.rev.value=Valid2(document.form.rev.value);

  if(document.form.rev.value==""){
    alert("Enter description");
    t=0;
   }  
  return;
}
function key(){
  document.form.key.value=Valid2(document.form.key.value);

  if(document.form.key.value==""){
    alert("Enter key words")
    t=0
  }
  return;
}
function sub(){
  t=1;
  
  log();
  key();
  rev();
  if(t == 1)
    form.submit(); 
}
</script>
</head>

<body>
<center>
  <img src="/pic2/vlinks.co.png"  alt="vlinks logo" title="Go to main page" 
         onclick="window.location='http://www.vlinks.co/vl_manager.asp' " >
  <h3>Management System v.1.1</h3>

<table border="1">
    <caption><strong>Add your site to DB</strong></caption>
    <%=Request.QueryString("msg")%>
  <form action="ins_url.asp" method=POST id=form name=form>
             
  <tr>  
     <td>Web site address :</td>
     <td><input type="text" name=ur size=30  maxlength="80" value="http://www." </td>
   </tr>
   
   <tr>
      <td>Keyword :</td>
      <td><textarea rows=7 cols=70 name=key maxlength="3000"  ></textarea></td>
    </tr>
    
     <tr>
       <td>Description :</td>
       <td><textarea rows=7 cols=70 name=rev maxlength="500"></textarea></td>
    </tr>
    
    <tr>
      <td>Icon's file name :</td>
      <td><input type="text"  name=pic size=40 maxlength="40" autofocus 
                placeholder="format: png, size: 50x150">
      </td>
     </tr>
    <tr>
      <td>Category :</td>
      <td>
        <SELECT  name=cat>
        <%do until rs2.EOF %>
        <OPTION value=<%=rs2("cat")%>><%=rs2("cat")%>
        </OPTION>
          <%
            rs2.MoveNext()
            loop
            %>
        </SELECT>
        
        <input type="button" value="Add new" style="color:green" onclick="window.location='new_cat.asp' ">
        </td>
      </tr>
   
     <tr>
        <td colspan=2>
             <input type="submit" value="Submit" style="width:100%; cursor:hand"  
                         onclick="sub()" id=button name=button>
           </form>
        </td>
     </tr>
  </table>
<br>
<p>
    <a href="http://www.w3.org/html/logo/">
          <img src="http://www.w3.org/html/logo/badge/html5-badge-h-css3.png" width="133" height="64" alt="HTML5 Powered with CSS3 / Styling" 
                  title="HTML5 Powered with CSS3 / Styling">
     </a> 
</p>
</center>
</body>

<%
rs2.Close
Set rs2=nothing
conn.Close
set conn=nothing
%>

</html>