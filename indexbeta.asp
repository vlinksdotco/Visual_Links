<%@ Language="VBScript" %>
<%
'
'   Copyright (C) 2017  Visual Links LTD.
'
'   The code in this page is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.


if Session("lan")="" then
 Session("lan")="in"
end if

Dim popular(25)
Dim temp(20)
Dim newweb(10)

check_id=0 

Session("lan")=Session("lan")
Dim i
i=0
Set conn = Server.CreateObject ("ADODB.connection")
conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; "  & _
"DBQ=" & Server.MapPath("db/main.mdb")

cat=Request.QueryString("cat")
SearchText=Request.QueryString("text1")

' count all web in table  --------------------------------------------
Set rs4 = Server.CreateObject ("ADODB.Recordset")

Set rs2 = Server.CreateObject ("ADODB.Recordset")
sql="select count(id) as urll from ur "
rs2.Open sql,conn
webcount=rs2("urll")
rs2.Close

'---------------------------------------------------------------------------
sql2="SELECT TOP 25 * FROM ur ORDER BY click DESC"
sql3="SELECT  * FROM ur ORDER BY id DESC"

rs2.Open sql2,conn

for i=0 to 24
popular(i)=rs2("id")
rs2.MoveNext 
next 
rs2.Close

rs2.Open sql3,conn

for i=0 to 9
newweb(i)=rs2("id")
rs2.MoveNext 
next 
rs2.Close

Function in_array(element)
	
rt=1

sql4="SELECT  * FROM ur "
  rs4.Open sql4,conn
   
    for k=0 to  element
       rs4.MoveNext 
       next
      
    temp_id=rs4("id")
      if temp_id="" then 
        rt=0
      end if

      if rt=1 then 
      for p=0 to 9
       if temp_id=popular(p) then 
          rt=0
          p=9
         end if  
      next 
   end if    
    
      if rt=1 then 
          for l=0 to 4
          if temp_id=newweb(l) then 
          rt=0
          l=4
         end if
       next       
       end if
        

rs4.close
		
If rt=0 then
in_array = True

Else 
in_array = false
check_id=temp_id
End If  

End Function

for i=0 to 9
Randomize
nRandom= Int((webcount*Rnd())+ 1)
    
    If in_array(nRandom) then
	i=i-1
    Else
	temp(i)=check_id
    End If 
 
next

' fill array and compare ---------------------------------------------------

Set rs2 = Server.CreateObject ("ADODB.Recordset")
sql="select count(id) as urll from ur where  rev like '%"&SearchText&"%' or key like '%"&SearchText&"%' or url like '%"&SearchText&"%' or cat like '%"&SearchText&"%'"
rs2.Open sql,conn
web_qty=rs2("urll")
rs2.Close
msg=Request.QueryString("msg")

if msg="" then 
  msg="" 
end if

cat=Request.QueryString("cat")

Set rs = Server.CreateObject ("ADODB.Recordset")
sql="select * from cat "


if  cat="" and SearchText="" then
sql2="SELECT TOP 25 * FROM ur ORDER BY click DESC"
sql3="SELECT  * FROM ur ORDER BY id DESC"

end if 

if not cat="" and SearchText="" then
sql2=" SELECT TOP 10000 * FROM ur where cat='"&cat&"' ORDER BY click DESC"
end if

if not SearchText=""  then
sql2=" SELECT TOP 1000 *  FROM ur  WHERE rev like '%"&SearchText&"%' or key like '%"&SearchText&"%' or url like '%"&SearchText&"%' or cat like '%"&SearchText&"%' ORDER BY [click] DESC "
end if


if not SearchText="" and not cat="" then
sql2=" SELECT TOP 10000 *  FROM ur  WHERE rev like '%"&SearchText&"%' or key like '%"&SearchText&"%' or url like '%"&SearchText&"%' and cat = '"&cat&"' ORDER BY [click] DESC "
end if

rs2.Open sql2,conn
rs.Open sql,conn
%>
<!DOCTYPE html> 
<html lang="en-US">                                                                            
<head>
<title>vlinks.co - Visual, Social & Anonymous Search Engine</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255"><!--utf-8"-->
<meta http-equiv="X-UA-Compatible" content="IE=edge"><meta name="MobileOptimized" content="320"><meta name="format-detection" content="telephone=no">
<meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes">
<meta name="HandheldFriendly" content="true">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black">
<meta name="apple-mobile-web-app-title" content="">
<meta name="description" content="Your Visual, Social & Anonymous Search Engine. Imagine no more text in your search results."> 
<meta name="keywords" content="vlinks, visual links, visual search engines, search visually, vlinks visual advertisement, visual advertising, no more text in your search results"> 
<meta name="google" content="notranslate"><meta name="robots" content="index, follow, archive">
<meta name="author" content="Visual Links Ltd.">

<link rel="sitemap" type="application/xml" title="Sitemap" href="/sitemap.xml">
<link rel="icon" href="/images/favicon.ico" type="image/x-icon">
<link rel="stylesheet"  type="text/css" href="/vendors/css/normalize.css">
<link rel="stylesheet"  type="text/css" href="/vendors/css/grid.css">
<link rel="stylesheet" href="http://www.w3schools.com/lib/w3.css">
<link rel="stylesheet"  type="text/css" href="/styles/desktop.css" media="screen and (min-width:481px)">
<link rel="stylesheet"  type="text/css" href="/styles/mobile.css" media="only screen and (max-width:480px)">
<link href='https://fonts.googleapis.com/css?family=Lato:400,100,300' rel='stylesheet' type='text/css'>
<link rel="canonical" href="http://vlinks.co/">

<script async type="text/javascript" src="/scripts/vlinks.js"></script>
<script async type="text/javascript" src="/scripts/mobile.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.3/jquery.min.js"></script>
</head>
<body style="position: relative; min-height: 100%; top: 0px;">
<div class="container">
<header>
  <!-- a href="index.asp">
    <img class="logo" src="pic2/vlinks.co.png" alt="VLinks's Logo" title="Show me main page" 
         style="left: 7px; top: 7px; position: absolute; border: 1; box-shadow: 3px 3px 2px #585858;">
  </a -->
  <h1>Your Visual, Social & Anonymous Search Engine.</h1>
  <h2>Imagine no more text in your search results.</h2>
 </header>

 <a href="index.asp">
    <img class="logo" src="pic2/vlinks.co.png" alt="VLinks's Logo" title="Show me main page">
  </a>

 <div id="input_bar" style="width:auto;"> 
  <div id="bar" style="height:23%; width:auto;">                                 
    <form method="get"  name="searchBox" autocomplete="on">&nbsp;&nbsp;&nbsp;
      <input style="height: 23px; font-family: 'Open Sans', sans-serif; font-weight: 300;" onfocus="bar" size="50"
             type="text" id=text1 name=text1 autofocus placeholder=" Search here...<%=msg%>" maxlength="80">&nbsp;
      <button type="button" onclick="document.getElementById('').innerHTML = submitForm()">&nbsp;Search&nbsp;</button>&nbsp;&nbsp;
      <img class="image_no_shadow" src="/images/plus.png" alt="Add your site/Report for broken link"
           title="Add your site / Report for broken link" 
           style="margin-top: -5px; position:absolute; height:auto; width:auto; cursor:pointer; border:1;"
           onclick="document.getElementById('').innerHTML = showse()">
    </form>
 </div>

<%if cat ="" and  SearchText="" then%>
<!-- ************************* Most popular rubric starts here *************************** -->
<div id="pop_rubric"> 
<h3>MOST POPULAR SITES</h3>&nbsp;
<% 
i=1
k=0
do until rs2.EOF or i>10
%>
              <a href="<%=rs2("url")%>" target="_blank">
                <img class="enlarge" src="pic2/<%=rs2("pic")%>" title="<%=rs2("rev")%>" alt="Popular Link"   
                     style="box-shadow: 3px 3px  2px #585858;"                                              onclick='aclick("openpage.asp?id=<%=rs2("id")%>")'>              </a>
 
<%if k=4 then%><br>&nbsp;<%end if%>
<%
rs2.MoveNext 
i=i+1
k=k+1
if k>4 then 
  k=0 
end if
loop
%>
</div>
<!-- *************************** Random rubric starts here *************************** -->
<div id="random_rubric">
<h3>TRY OUR RANDOM SITES</h3>&nbsp;
<%
    k=0

    for i=0 to 9
    sql4="SELECT  * FROM ur where id="&temp(i)
    rs4.Open sql4,conn
%>
    
          <a href="<%=rs4("url")%>" target="_blank">
              <img class="enlarge"src="pic2/<%=rs4("pic")%>"  title="<%=rs4("rev")%>" alt="Random Link"   
                   style="box-shadow: 3px 3px 2px #585858;"                                            onclick='aclick("openpage.asp?id=<%=rs4("id")%>")'>          </a>
     
<%if k=4 then %><%k=-1%><br>&nbsp;<%end if%>
<%  
  rs4.close
   k=k+1
  next  
%>
</div>
<!-- *************************** New coming rubric starts here **************************-->
<div id="new_rubric">
<h3>CHECK OUR NEW SITES</h3>&nbsp;
<%
rs2.Close 
rs2.Open sql3,conn
%>
<% 
i=1
k=0
do until rs2.EOF or i>5
%>
        <a href="<%=rs2("url")%>" target="_blank">
          <img class="enlarge" src="pic2/<%=rs2("pic")%>" title="<%=rs2("rev")%>" alt="New Link"  
               style="box-shadow: 3px 3px 2px #585858;"                                  onclick='aclick("openpage.asp?id=<%=rs2("id")%>")'>         </a>

<%if k=4 then%><!--br-->&nbsp;<%end if%>
<%
rs2.MoveNext 
i=i+1
k=k+1
if k>4 then k=0 end if
loop
%>	
<%end if%>
</div>
<!-- by cat -->&nbsp;
<%
 if not cat="" and SearchText="" then
 k=0
 do until rs2.EOF %>&nbsp;
 <a href="<%=rs2("url")%>" target="_blank">
   <img class="enlarge" src="pic2/<%=rs2("pic")%>" title="<%=rs2("rev")%>"  alt="category" 
        style="box-shadow: 3px 3px 2px #585858;"        onclick='aclick("openpage.asp?id=<%=rs2("id")%>")'>
 </a>

<%if k=4 then%><br>&nbsp;<%end if%>
<%
rs2.MoveNext 
k=k+1
if k>4 then k=0 end if
loop
%>    
<%end if%>
<!-- by search1 -->
<% if not SearchText="" and cat="" then %>
<!-- ************************* Advertisers rubric starts here ************************* -->
<div id="ads_rubric" style="margin-left:-4px;">
<h4>OUR ADVERTISERS</h4>
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="a1">
      <img class="enlarge" src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858" 
        onclick='aclick("openpage.asp?id=2625")' id="a2"></a>&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="b1">
      <img class="enlarge" src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858"
        onclick='aclick("openpage.asp?id=2625")' id="b2"></a>&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="c1">
      <img class="enlarge" src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858"
        onclick='aclick("openpage.asp?id=2625")' id="c2"></a>&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="d1">
      <img class="enlarge" src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858"
        onclick='aclick("openpage.asp?id=2625")' id="d2"></a>&nbsp;
  <a href="http://www.vlinks.co/vlinks_ads.asp" target="_blank" id="e1">
      <img class="enlarge" src="images/adver.png" alt="ad website" 
        title="Your website can be here too !" style="box-shadow: 3px 3px 2px #585858"
        onclick='aclick("openpage.asp?id=2625")' id="e2"></a>
</div>
<br>
<%if rs2.EOF then %>  
    <audio autoplay="autoplay">  <source src="sounds/Banana_Slap.mp3" type="audio/mpeg">
    </audio>
    <h3 style="text-align:center; color:red;">No results found.</h3>        <!-- Show where else to search -->
    <p style="text-align:center; color:#009933; text-shadow: 1px 1px 3px #585858;">       <h3>Try to find below OR press on orange plus icon to add your web site.</h3>    </p>    <br>
       <a href="https://www.google.co.il/?gfe_rd=cr&ei=zMXuVabjIemT8QfXgZ-wCA&gws_rd=cr#q=<%=SearchText%>" target="_blank">
          <img class="enlarge" src="pic2/google.com.png" alt="Google" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="https://en.wikipedia.org/w/index.php?title=Special:Search&profile=default&fulltext=Search&search=<%=SearchText%>&searchToken=di3txgtu4gp67p4vojfj6bf8x" target="_blank">
          <img class="enlarge" src="pic2/wikipedia.org.png" alt="wikipedia.org" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="http://www.bing.com/search?q=<%=SearchText%>&qs=n&form=QBLH&pq=<%=SearchText%>&sc=0-0&sp=-1&sk=&cvid=dd14addde2474605a2fb1b5726831246&adlt=strict" target="_blank">
          <img class="enlarge" src="pic2/bing.com.png" alt="Bing" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="https://www.youtube.com/results?search_query=<%=SearchText%>" target="_blank">
          <img class="enlarge" src="pic2/youtube.com.png" alt="YouTube" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="https://search.yahoo.com/search;_ylt=A2KJyw5dJflV03gBOzabvZx4?p=<%=SearchText%>&toggle=1&cop=mss&ei=UTF-8&fr=yfp-t-901&fp=1" target="_blank">
          <img class="enlarge" src="pic2/yahoo.com.png" alt="Yahoo" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>       <br>
       <a href="https://duckduckgo.com/?q=<%=SearchText%>" target="_blank">
          <img class="enlarge" src="pic2/duckduckgo.com.png" alt="DuckDuckGo" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="https://www.yandex.com/search/?text=<%=SearchText%>" target="_blank">
          <img class="enlarge" src="pic2/yandex.com.png" alt="Yandex" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="https://vimeo.com/search?q=<%=SearchText%>" target="_blank">
          <img class="enlarge" src="pic2/vimeo.com.png" alt="Rambler" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="https://translate.google.com/#auto/en/<%=SearchText%>" target="_blank">
          <img class="enlarge" src="pic2/translate.google.com.png" alt="Google Translate" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
       <a href="http://www.ask.com/web?q=<%=SearchText%>&qsrc=0&o=0&l=dir&qo=homepageSearchBox" target="_blank">
          <img class="enlarge" src="pic2/ask.com.png" alt="Ask" title=""
               style="box-shadow: 3px 3px 2px #585858;">
       </a>
    </p>      
<%else%><h3 style="text-align:center; color:#31B404;">Found <%=web_qty%> results:</h3>
<% end if %>

<% 
i=0
k=0
do until rs2.EOF 
%>
  <a href="<%=rs2("url")%>" target="_blank">
    <img class="enlarge" src="pic2/<%=rs2("pic")%>" title="<%=rs2("rev")%>" alt="Search Result" border="1" 
         style="box-shadow: 3px 3px 2px #585858; cursor:pointer; cursor:hand"         onclick='aclick("openpage.asp?id=<%=rs2("id")%>")'>
  </a>

<%if k=4 then%><br>&nbsp;<%end if%>

<%
rs2.MoveNext 
i=i+1
k=k+1
if k>4 then k=0 end if
loop
%>	
<%end if%>
<!-- by search2 -->&nbsp;
<% if not SearchText="" and not cat="" then %>
<%
sql2="SELECT TOP 800 *  FROM ur  WHERE cat='"&cat&"' and (rev like '%"&SearchText&"%' or key like '%"&SearchText&"%' or url like '%"&SearchText&"%') ORDER BY [click] DESC "
k=0
rs2.Close 
rs2.Open sql2,conn
 do until rs2.EOF %>

  <a href="<%=rs2("url")%>" target="_blank">
    <img class="enlarge" src="pic2/<%=rs2("pic")%>" title="<%=rs2("rev")%>" alt="Search Result" border="1" 
         style="box-shadow: 3px 3px 2px #585858; cursor:pointer; cursor:hand" 
         /*onmouseover="window.status='<%=rs2("url")%>'" onmouseout="window.status=''" */          onclick='aclick("openpage.asp?id=<%=rs2("id")%>")'>
  </a>

<%if k=4 then%><br>&nbsp;<!--tr-->
<%end if%>

<%
rs2.MoveNext 

k=k+1
if k>4 then k=0 end if
loop
%>
<%end if%>

 <footer>
<!-- p>
        <a href="/">Home</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
        <a href="/about.html">About</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
        <a href="/vlinks_ads.asp">Advertising</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
        <a href="/privacy_policy.html">Privacy</a>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;
        <a href="/terms.html">Terms</a>
     </p -->
 <span>&copy;&nbsp;2017&nbsp;Visual Links Ltd.</span>
      All rights reserved.&nbsp;The website is not responsible for any linked web site.
<a href="http://www.vlinks.co" title="Back to top page." class="tooltip">
    <img alt="Home" src="http://www.vlinks.co/images/favicon.ico" style="height:auto;width:auto;cursor:pointer; border:0"></a>&nbsp;&nbsp;
<!-- Social presence -->
<a href="mailto:info@vlinks.co" title="Email to us." target="_blank" class="tooltip">
    <img alt="Email" src="http://www.vlinks.co/images/mail.png" style="height:auto;width:auto;cursor:pointer; border:0"></a>&nbsp;&nbsp;
<a href="https://www.facebook.com/pages/Visual-Links-vlinksco/183003611893479" title="Find us on Facebook" target="_blank" class="tooltip">
    <img alt="Facebook" src="http://www.vlinks.co/images/facebook.png" style="height:auto;width:auto;cursor:pointer; border:0"></a>&nbsp;&nbsp;
<a href="https://twitter.com/VisualLinks" title="Find us on Twitter" target="_blank" class="tooltip">
    <img alt="Twitter" src="http://www.vlinks.co/images/twitter.png" style="height:auto;width:auto;cursor:pointer; border:0"></a><br>
 </footer>

<div id=pa1 style="visibility:hidden; width:5; height:5" ><!--iframe src="clon.asp"  height=5 width=5></iframe-->  
</div>
<!-- *************************** Draw plus window starts here *************************** -->
<div id=se style="text-align:center; position:absolute; left:20%; top:10%; 
                  font-size:16px; font-weight:200; background-color:white; 
                  border-style:solid; border-color:orange; border-width:2px; z-index:2; visibility:hidden;">
 <form action="/ins_mas.asp" method=POST id=form2 name=form2><br>
  <div style="text-align:center;">
  <label>WWW:</label>
  <input type="text" name="ur" size="30" placeholder="Enter here your web site / Report for problems" maxlength="140">
  </div>
  <div style="text-align:center;">
  <label>Describe:</label>
    <input type="text" name="mas" size="30" placeholder="(Optional) Enter here ..." maxlength="140">
    <input type="button" value="Send" style="width:23%; border-width:2px;" id=button name=button2 onclick="flag2()">
    <input type="button" value="Cancel" style="width:23%; border-width:2px;" id=button2 onclick="hidese()">
  </div>
 </form>
</div>
<script>
  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');

  ga('create', 'UA-43570893-1', 'vlinks.co');
  ga('send', 'pageview');
</script>
</div>
   </div>
  </body>
</html>

<%
rs.Close
Set rs=nothing
rs2.Close
Set rs2=nothing

conn.Close
set conn=nothing
%>