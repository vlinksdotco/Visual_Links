<%@ Language=VBScript %>

<!doctype html>
<html lang="en">
<head>
<meta http-equiv="content-type" content="text/html; charset=windows-1255">  <!-- charset=utf-8">  -->
<title>New category</title>
<link rel="stylesheet" href="desktop.css">
</head>
<body><center>
<img src="/pic2/vlinks.co.png"  alt="vlinks logo" title="Go to main page"          onclick="window.location='http://www.vlinks.co/vl_manager.asp' " >
<h1>Management System v. 1.1</h1>
<hr>
<script language="javascript">
t=1

//
function Valid2(str){
  r=""
  var newstr="";
  var a,count;
  a=str.split("'");
  count=a.length;
  for(i=0;i<count;i++)
    newstr=newstr+a[i];
		
  a=newstr.split(",");
  count=a.length;
  for(i=0;i<count;i++)
    r=r+a[i];
		
  return r ;
}

//
function log(){
  document.form.cat.value=Valid2(document.form.cat.value)

  if(document.form.cat.value==""){
    alert("Enter cat")
    t=0
  } 
  return;
}

//
function sub(){
 t=1
 log()
  
 if(t==1)
   form.submit() 
}
</script>
<body><center>
 <table>
 <font color=red><%=Request.QueryString("msg")%>  </font>

 <tr><td colspan=2 align=center><%=Request.QueryString("msg")%> </td>
    <tr>
    <FORM action="ins_cat.asp" method=POST id=form name=form>
   <td>Create a new category:</td><td><INPUT type="text"  name=cat></td>
  <tr>
    <td colspan=2><INPUT type="button" value="Submit" style="width:100%;cursor:hand" 
                   onclick=sub() id=button1 name=button1></td>
 </table>
</form>
</center>
</body>
</html>
