<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
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
    padding:30px;	 	 
}
#footer {
    background-color:#009;
    color:yellow;
    clear:both;
    text-align:center;
   	padding:5px;	 	 
}
.button { 
	width: 200px; 
	height : 60px; 
	text-align: center;
} 
table {
    border-collapse: collapse;
    border: 1px solid #009;
}
a {
    color: black;
	text-decoration: none;
	background-color:#FFF;
}
.box {
    width: 150px;
    height: 45px;
    border: 3px solid black;
	text-align:center;
	color: black;
	margin: auto;
	font-size:20px
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Update Order View</title>
</head>

<script type="text/javascript">
//This function is declared in order to validate the Username and Password fields.
function validateChequeNumber() {
	//These variables are declared by getting the data from the form labelled ID.
	var ChequeNumber = document.getElementById("ChequeNumber").value;
	
	//This if statement is going to alert the user with an error message if the ChequeNumber is left empty.
	if (ChequeNumber == "") 
		{
		alert("ChequeNumber is required");
		return false;
		}
	//This if statement below is to check if the value entered in the ChequeNumber field is a number.
	else
		{
		if (/^\d+(\.\d{0})?$/.test(ChequeNumber)==false)
		{
		alert("ChequeNumber must be a number.");
		return false
		}	
		}
}
</script>

<%
'If the form is not empty then the system can do the following.
If request.form <> "" Then
	Dim Con, rs, sql, validateChequeNumber, ChequeNumber
		
	Set Con = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
		
	Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))
		
	'This SQL statement is selecting the ChequeNumber from the table tblOrder where the ChequeNumber is equal to the ChequeNumber Finance has inputted.
	sql = "SELECT ChequeNumber FROM tblOrder WHERE ChequeNumber = " & request.form("ChequeNumber")
	rs.Open sql, Con
		
	'This if statement goes through all the ChequeNumber in the table tblOrder and if the end of file is false then that means the cheque number Finance has inputed already exists.
	If rs.EOF = False Then
		validateChequeNumber = "This cheque number already exists."
	End if
	
	'This if statement checks if the variable validateChequeNumber is empty meaning the cheque number does not exists then the system will do the following.
	If validateChequeNumber = "" Then
		'This SQL statement will update the table tblOrder by setting the ChequeNumber to the ChequeNumber Finance has entered where the OrderID is equal to the OrderID for that order.
		sql = "UPDATE tblOrder SET ChequeNumber = '"&request.form("ChequeNumber")&"' WHERE OrderID = " & request.QueryString("OrderID")
		con.execute(sql)
		con.close	
		'The system will then redirect to the OrderView page where it will show the final order.
		response.redirect("OrderViewF.asp?OrderID="&request.QueryString("OrderID")&"")
	End if

End if

%>

<body>
<div id="header">
<div align="center">
  <table width="419" align="center">
    <tr>
      <td width="127" height="60" align="center"><strong><a href="logout.asp">LOG OUT</a></strong></td>
      <th width="280" rowspan="5"><img src="http://www.jobs.ac.uk/images/employer-logos/medium/1164.gif" alt="l" width="152" height="124" align="left" /></th>
      </tr>
    <tr>
    </tr>
  </table>
</div>
<h1>ORDER SYSTEM</h1>
</div>

<div id="section">
<div align="center">
  <table width="126" border="1">
    <tr>
      <td width="169">
        <div class="box"> 
        <div align="center"><a href="displayFinanceHomepage.asp"><strong>HOMEPAGE</strong></a>
        </div>
      </td>
    </tr>
  </table>
</div>

<p align="center">----------------------------------------------------------------------------------------------------</p>

<h2 align="center"><u><strong>SUBMIT ORDER</strong></u></h2>

<div align="centre">
	<%=validateChequeNumber%>
</div>
  
  <form name="ChequeNumber" onsubmit="return validateChequeNumber(this)" method="post" action="updateOrderView.asp?OrderID=<%=request.QueryString("OrderID")%>">
  <p align="center"><strong>Cheque number: </strong>
    <input name="ChequeNumber" type="text" id="ChequeNumber"/>
  </p> 
   <div align="center">
        <p>
          <input name="Submit" type="submit" value="FINALISE ORDER">
        </p>
	</div>
</form>

</div>

<div id="footer">
Copyright © LSC
</div>

</body>
</html>
