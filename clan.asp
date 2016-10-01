<%@ Language=VBScript %>
<%
lan=Request.QueryString("lan")
Session("lan")=lan  
Response.Redirect "index.asp"
%>