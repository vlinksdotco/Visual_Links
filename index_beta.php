<?php
// print before html code - echo nl2br("Hello world from PHP !\n");
?>

<!DOCTYPE html>
<html>
<title>vlinks.co v.2</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="http://www.w3schools.com/lib/w3.css">
<link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Raleway">
<link rel="stylesheet"  type="text/css" href="/styles/desktop.css" media="screen and (min-width:481px)">
<style>
body,h1 {font-family: "Raleway", sans-serif}
body, html {height: 100%}
.bgimg {
    background-image: url('http://www.w3schools.com/w3images/forestbridge.jpg');
    min-height: 100%;
    background-position: center;
    background-size: cover;
}
</style>
<body>
<div class="bgimg w3-display-container w3-animate-opacity w3-text-white">
  <div class="w3-display-topleft w3-padding-large w3-xlarge">
    <img id="logo" src="pic2/vlinks.co.png"   alt="Visual Links's Logo" title="Show me main page" 
         style="left:7px;top:7px; position:absolute; border:1; box-shadow: 3px 3px 2px #585858;"     
         onclick="window.location='index.asp'">

  <h2>Your Visual, Social & Anonymous Search Engine.</h2>
  <h3>Imagine no more text in your search results.</h3>
</div>

  <div class="w3-display-middle">
    <h1 class="w3-jumbo w3-animate-top">COMING SOON</h1>
    <hr class="w3-border-grey" style="margin:auto;width:40%">
    <p class="w3-large w3-center">35 days left</p>
  </div>
</div>

<?php
/*
 *  filename : test.php
 */

   // 
	echo nl2br("Printing php code from inside html body Hello world from PHP !\n");
      echo "Hello PHP world!<br>";
	
	// sum of 2 nums
	$num1=4;
	$num2=3;
	$sum=$num1+$num2;
	
	echo "sum = $sum<br>";
	echo "<br>End of php."
?>

</body>
</html>