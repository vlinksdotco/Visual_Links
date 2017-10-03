<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
t=1
function Valid2(str)
{
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

function log()
{

document.form.cat.value=Valid2(document.form.cat.value)

  if(document.form.cat.value=="")
   {
    alert("Enter cat")
    t=0
   }
   
    return;
}

function sub()
{
 t=1
 log()
  
if(t==1)
 form.submit() 
}
</script>

<BODY>
 <table>
 <font color=red><%=Request.QueryString("msg")%> </font>

 <tr><td colspan=2 align=center><%=Request.QueryString("msg")%> </td>
    <tr>
    <FORM action="ins_catgam.asp" method=POST id=form name=form>

    <td>category</td><td><INPUT type="text"  name=cat></td>
  <tr>
    <td colspan=2><INPUT type="button" value="Submit" style="width:100%;cursor:hand" onclick=sub() id=button1 name=button1></td>
 </table>
</FORM>

</BODY>
</HTML>