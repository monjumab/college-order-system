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
<title>Update Order</title>
</head>

<script type="text/javascript">
//This function is declared in order to validate the ProductID and Quantity fields if they are left empty and ensures that the Quantity is a number between 1 and 50.
function validateEditForm() {
	//These variables are declared by getting the data from the form labelled ID.
	var ProductID = document.getElementById("ProductID").value;
	var Quantity = document.getElementById("Quantity").value;

	//This if statement is going to alert the user with an error message if both the ProductID and Quantity fields are left empty.
	if (ProductID == "" && Quantity == "") 
		{
		alert( "ProductID and Quantity is required.");
		return false;
		}
	//This if statement is going to alert the user with an error message if just the ProductID is left empty.
	if (ProductID == "") 
		{
		alert("ProductID is required");
		return false;
		}
	//This if statement is going to alert the user with an error message if just the Quantity is left empty. If the Quantity is not left empty then it will check whether the value the teacher has entered is a number or not and then will check if it is a number between 1 and 50.		
	if (Quantity == "") 
		{
		alert("Quantity is required");
		return false;
		}
	else
		{
		if (/^\d+(\.\d{0})?$/.test(Quantity)==false)
		{
		alert("Quantity must be a number.");
		return false
		}
		if (Quantity <1 || Quantity > 50)
		{
		alert("Quantity must be between 1 and 50")
		return false	
		}	
		}
		
}

</script>

<%
dim Con, rs, sql, errorMessage

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is selecting everything from the table tblOrderDetail where the OrderDetailID is equal to the OrderDetailId from the tblOrderDetail.
sql = "SELECT * FROM tblOrderDetail WHERE OrderDetailID = " & request.QueryString("OrderDetailID") 

rs.Open sql, Con

	'This if statement is checking if the form is not empty. If it is not empty then it will go to the next if statement underneath.
	If request.form <> "" Then
		'This if statement is checking if the ProductID or the Quantity is left empty and if it is then it will display an error message.
		If request.form("ProductID") = "" OR request.form("Quantity") = "" Then
			errorMessage = "Please fill in the fields."
		Else
		'If the ProductID or Quantity is not left empty then this SQL statement will update the tblOrderDetail and update the ProductID and the Quantity into the values the teacher has inputed.
			sql = "UPDATE tblOrderDetail SET ProductID = '"&request.form("ProductID")&"', Quantity = '"&request.form("Quantity")&"' WHERE OrderDetailID = " & request.QueryString("OrderDetailID")
			con.execute(sql)
			'The system will then redirect to the edit page.
			response.redirect("editOrder.asp?OrderID=("&rs("OrderID")&")")
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

<div align="right">
<!--This displays the current date.-->
  <%=Date()%>
</div>

<div align="center">
   <table width="179" border="1">
    <tr>
      <td width="169">
        <div class="box"> 
        <div align="center"><a href="displayTeacherHomepage.asp"><strong>HOMEPAGE</strong></a>
        </div>
        </td>
        <td width="169">
        <div class="box"> 
        <div align="center"><a href="editOrder.asp?OrderID=<%=rs("OrderID")%>"><strong>EDIT ORDER</strong></a>
        </div>
        </td>
      </tr>
  </table>
  <p>----------------------------------------------------------------------------------------------------</p>
</div>

<h2 align="center"><strong><u>UPDATE ORDER</u></strong></h2>

  <form name="updateOrder" onsubmit="return validateEditForm(this)" method="post" action="updateOrder.asp?OrderDetailID=<%=rs("OrderDetailID")%>">
  <!--If the error message from the above code is needed to display, this will display the error message to the user on the page.-->
  <div align="center"><h4><%=errorMessage%></h4></div>
  <p align="center"><strong>Product ID:</strong>
    <!--The asp is requesting the ProductID and displaying it in the textfield to show the teacher what their original ProductID was.-->
    <input name="ProductID" type="text" id="ProductID" value="<%=rs("ProductID")%>"/> 
     <strong>Quantity: </strong>
     <!--The asp is requesting the Quantity and displaying it in the textfield to show the teacher what their original Quantity was.-->
     <input name="Quantity" type="text" id="Quantity" value="<%=rs("Quantity")%>"/>
  </p>
  <p align="center">&nbsp;</p>
  
    <div align="center">
        <p>
          <input name="Submit" type="submit" value="UPDATE">
        </p>
	</div>
  
</form>

<div id="footer">
Copyright © LSC
</div>

</body>
</html>
