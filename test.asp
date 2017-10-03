<!DOCTYPE HTML><html dir="ltr" lang="en-US">	<head>	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />	<title>test.asp</title>	<!-- CSS -->	<style type="text/css">
	body{ font-size:18px; }
	form{text-align:center; width:800px; margin:300px auto;}
	  #search{
		width:600px;
		padding:8px 15px;
		border:1px solid blue;
	  }
	  #button{
	   position:relative;
		padding:8px 15px;
		background-color:orange;
		border:1px solid blue;
		color:blue;
		/*margin-left: -5px;*/
		cursor:pointer;
   	  }
     #button-hover{
		background-color:red;
		transition:all 0.40s;
    }  	
	</style>	</head>	<body>			<form action="https://google.com">				<input  style="font-size:18px;" type="text" placeholder="Search..." maxlength="50" id="search">			    <input style="font-size:18px;" type="submit" value="Search" id="button">			</form>			</body></html>