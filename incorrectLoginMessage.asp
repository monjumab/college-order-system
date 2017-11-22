<!DOCTYPE html>
<html>
<head>
<style>
#header {
    background-color:#009;
    color:yellow;
    text-align:center;
    padding:5px;
}
#section {
	text-align:center;
    padding:50px;	 	 
}
#footer {
    background-color:#009;
    color:yellow;
    clear:both;
    text-align:center;
   padding:5px;	 	 
}
</style>

<script type="text/javascript">
//This function is declared in order to validate the Username and Password fields if they are left empty.
function validateLoginForm() {
	//These variables are declared by getting the data from the form labelled ID.
	var Username = document.getElementById("Username").value;
	var Password = document.getElementById("Password").value;
	
	//This if statement is going to alert the user with an error message if both the Username and Password fields are left empty.
	if (Username == "" && Password == "") 
		{
		alert( "Username and password is required.");
		return false;
		}
	//This if statement is going to alert the user with an error message if just the Username is left empty.
	if (Username == "") 
		{
		alert("Username is required");
		return false;
		}
	//This if statement is going to alert the user with an error message if just the Password is left empty.
	if (Password == "") 
		{
		alert("Password is required");
		return false;
		}
}
</script>

</head>
<body>

<div id="header"><img src="http://www.jobs.ac.uk/images/employer-logos/medium/1164.gif" alt="l" width="152" height="124" align="center" />
<h1>ORDER SYSTEM</h1>
</div>

<div id="section">
<h2>LOGIN</h2>
<p style="color:red"><strong>Login details are incorrect. Please try again.</strong></p>
<form name="Login" onsubmit="return validateLoginForm(this)" method="post" action="login.asp">
  <strong>Username:</strong>
  <input name="Username" type="text" id="Username" placeholder="Type your username">
  <br>
  <br>
  <strong>Password:</strong>
  <input name="Password" type="password" id="Password" placeholder="Type your password">
  <br>
  <br>
  <div id="submit">
  <input name="Sign in" type="submit" value="Sign in">
</div>
</form>

</div>

<div id="footer">
Copyright © LSC
</div>

</body>
</html>