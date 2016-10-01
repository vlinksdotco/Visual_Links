<%@ LANGUAGE="VBSCRIPT" %>

<%
     ' Connects and opens the text file
     ' DATA FORMAT IN TEXT FILE= "username<SPACE>password"

     Set MyFileObject=Server.CreateObject("Scripting.FileSystemObject")
     Set MyTextFile=MyFileObject.OpenTextFile(Server.MapPath("\passwords123.txt"))

     ' Scan the text file to determine if the user is legal
     WHILE NOT MyTextFile.AtEndOfStream
              ' If username and password found
          	IF MyTextFile.ReadLine = Request.form("username") & " " & Request.form("password") THEN
               	' Close the text file
               	MyTextFile.Close
               	' Go to login success page
               	Session("GoBack")=Request.ServerVariables("SCRIPT_NAME")
               	Response.Redirect "vl_manager.asp"
               	Response.end
          	END IF
     WEND

     ' Close the text file
     MyTextFile.Close
     ' Go to error page if login unsuccessful
     Session("GoBack")=Request.ServerVariables("SCRIPT_NAME")
     Response.Redirect "login.asp"
     Response.end

%>