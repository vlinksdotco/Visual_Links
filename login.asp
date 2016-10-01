<!doctype html>
<html>
<head><title>Login</title>
<meta http-equiv="content-type" content="text/html;  charset=utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0"> 
<meta name="robots" content="noindex">
<link rel="stylesheet"  type="text/css"
         href="/styles/mobile.css" media="only screen and (max-width: 480px)">
<link rel="stylesheet"  type="text/css"
         href="/styles/desktop.css" media="screen and (min-width: 481px )">
<!-- [if IE]>
<link rel="stylesheet" type="text/css" href=explorer.css" media="all />
<! [endif] -->
</head>

<body style="color:green; background-color:black;">
<div style=" background-color:black; color:white; margin:20px; padding:20px;">
<form style="display:block;margin-top: 0em;" method="post" action="engine.asp">
<fieldset>
  <legend style="color:green;
    display: block;
    padding-left: 2px;
    padding-right: 2px;
    border:none;">Log In
  </legend>
<font color="green">
Username:  <input type="text" name="username" id="Text1" value="" size="8" maxlength="20" autofocus><br><br>
Password:  <input type="password" name="password" id="Password1" value="" size="20" maxlength="20"><br><br>
</font>
<button type="submit">Log in</button>
</fieldset>
</form>
</div>
</body>
</html>