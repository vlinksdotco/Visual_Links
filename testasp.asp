<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<body>
<%
response.write("<h2 style='color:red'>ASP test ...</h2>")
%>

<p>VBScript - default scripting language in ASP</p>

<%
dim name
name="Visual Links Ltd."
response.write(name & "&nbsp")
'changing value
name="vlinks.co"
response.write(name & "<br>")

dim i, j
For j = 0 To 10
 response.write(j & "&nbsp")
Next
 response.write("<br>")

'array
Dim famname(5)
famname(0) = "Jan Egil"
famname(1) = "Tove"
famname(2) = "Hege"
famname(3) = "Stale"
famname(4) = "Kai Jim"
famname(5) = "Borge" '??? six elements

For i = 0 to 5
 response.write(famname(i) & "<br>")
Next

'date
Function myfunction()
 myfunction=Date()
End Function

response.write("Date : " & myfunction() & "&nbsp")

i=hour(time)
If i < 10 Then
   response.write("Good morning!" & "&nbsp")
Else
   response.write("Have a nice day!" & "&nbsp")
End If

'print what day
d=weekday(Date)

Select Case d
  Case 1
    response.write("Sleepy Sunday")
  Case 2
    response.write("Monday again!")
  Case 3
    response.write("Just Tuesday!")
  Case 4
    response.write("Wednesday!")
  Case 5
    response.write("Thursday...")
  Case 6
    response.write("Finally Friday!")
  Case Else
    response.write("Super Saturday!!!!")
End Select

%>

<form method="post" action=""><!--"simpleform.asp"-->
<p>
First Name: <input type="text" name="fname"><br><br>
Last Name: <input type="text" name="lname"><br><br>
<input type="submit" value="Submit">
</p>
</form>

<%
response.write(request.form("fname"))
response.write(" " & request.form("lname"))
%>

</body>
</html>