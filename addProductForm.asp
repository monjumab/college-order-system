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
a {
    color: black;
	text-decoration: none;
	background-color:#FFF;
}
table {
    border-collapse: collapse;
    border: 1px solid #009;
}
.box {
	width: 140px;
    height: 45px;
    border: 3px solid black;
	text-align:center;
	color: black;
	margin: auto;
	font-size:20px
}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Products</title>
</head>

<script type="text/javascript">
//This function is declared in order to validate the ProductID, LegerCode, SupplierName, Tele, Code, Description, UnitCost and ProductVAT.
function validateProductForm() {
	//These variables are declared by getting the data from the form labelled ID.
	var ProductID = document.getElementById("ProductID").value;
	var LegerCode = document.getElementById("LegerCode").value;
	var SupplierName = document.getElementById("SupplierName").value;
	var Tele = document.getElementById("Tele").value;
	var Code = document.getElementById("Code").value;
	var Description = document.getElementById("Description").value;
	var UnitCost = document.getElementById("UnitCost").value;
	var ProductVAT = document.getElementById("ProductVAT").value;
		
	//These if statements below are going to alert the user with an error message if all the required fields are left empty.
	if (ProductID == "" && LegerCode == "" && SupplierName == "" && Tele == "" && Description == "" && UnitCost == "") 
		{
		alert( "The following fields are required: ProductID, LegerCode, SupplierName, Tele, Description and UnitCost.");
		return false;
		}
	if (LegerCode == "" && SupplierName == "" && Tele == "" && Description == "" && UnitCost == "") 
		{
		alert( "The following fields are required: LegerCode, SupplierName, Tele, Description and UnitCost.");
		return false;
		}
	if (SupplierName == "" && Tele == "" && Description == "" && UnitCost == "") 
		{
		alert( "The following fields are required: SupplierName, Tele, Description and UnitCost.");
		return false;
		}
	if (Tele == "" && Description == "" && UnitCost == "") 
		{
		alert( "The following fields are required: Tele, Description and UnitCost.");
		return false;
		}
	if (Description == "" && UnitCost == "") 
		{
		alert( "The following fields are required: Description and UnitCost.");
		return false;
		}
							
	if (ProductID == "") 
		{
		alert("ProductID is required");
		return false;
		}
		
	if (LegerCode == "") 
		{
		alert("LegerCode is required");
		return false;
		}
	//This is checking whether the value in LegerCode is a number.
	else
		{
		if (/^\d+(\.\d{0})?$/.test(LegerCode)==false)
		{
		alert("LegerCode must contain a number. Check the table on the left.");
		return false
		}	
		}
	
	if (SupplierName == "") 
		{
		alert("SupplierName is required");
		return false;
		}
	else
	//This is checking if the SupplierName has more than 4 characters and less than 21 characters otherwise the data won't be accepted.
		{
		if (SupplierName.length < 5 || SupplierName.length > 20)	
		{
		alert("SupplierName has to be more than 4 characters and less than 21 characters.")
		return false
		}
		}
		
	if (Tele == "") 
		{
		alert("Tele is required");
		return false;
		}
	//This is checking whether the value in Tele is a number.
	else
		{
		if (/^\d+(\.\d{0})?$/.test(Tele)==false)
		{
		alert("Tele must contain a number.");
		return false
		}
		}
	//This if statement is checking if the telephone number is 11 digits and nothing more, nothing less than that.
	if (Tele != "")
		{
		if(Tele.length <11 || Tele.length >11)
		{
		alert("Telephone number must be 11 digits.")
		return false
		}
		}	
		
	if (Description == "") 
		{
		alert("Description is required");
		return false;
		}
	else
	//This is checking if the Description has more than 4 characters and less than 151 characters otherwise the data won't be accepted.
		{
		if (Description.length < 5 || Description.length > 150)	
		{
		alert("Description has to be more than 4 characters and less than 151 characters.")
		return false
		}
		}
			
	if (UnitCost == "") 
		{
		alert("UnitCost is required");
		return false;
		}
	//This is checking whether the value in UnitCost is a number or a decimal that is up to 2 decimal points.
	else
		{
		if (/^\d+(\.\d{2})?$/.test(UnitCost)==false)
		{
		alert("UnitCost must be a number or a decimal up to 2 decimal points.");
		return false
		}	
		}		
	//This is checking whether the user inputed 0 in the UnitCost field and if it is then it won't be accepted.
	if (UnitCost == "0")
	{
	alert("Unit Cost cannot be £0.")
	return false
	}

	//This is checking whether the value in ProductVAT is a number or a decimal that is up to 2 decimal points.		
	if (ProductVAT != "")
		{
		if (/^\d+(\.\d{2})?$/.test(ProductVAT)==false)
		{
		alert("ProductVAT must be a number or a decimal up to 2 decimal points.");
		return false
		}			
		}	
	//This is checking whether the user inputed 0 in the ProductVAT field and if it is then it won't be accepted.
	if (ProductVAT == "0")
	{
	alert("ProductVAT cannot be £0.")
	return false
	}	
}

</script>

<%
'If the form is not empty then the system will do the following.
If request.form <> "" Then

	dim Con, rs, sql, validateProduct, pletter, pletterError, ProductID, LegerCode, validateLegerCode, LegerCodeExist
	
	Set Con = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	
	Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))
	
	'This SQL statement is selecting the ProductID's from the table tblProduct.
	sql = "SELECT ProductID FROM tblProduct"
	
	rs.Open sql, Con
	
	'The ProductID that has been entered by the user will be assigned to the variable ProductID.
	ProductID = request.form("ProductID")
	
	'The left most letter of the ProductID is assigned to the variable pletter.
	pletter = left(ProductID, 1)
	
	'This if statement is checking whether the variable pletter is equal to "P" and if it is then the system will do the following.
	if pletter = "P" Then
	
		'Assign nothing to the variable validateProduct.
		validateProduct = ""
		
		'This is a loop that will not end until end of file.
		Do while not rs.EOF
		'This if statement is seeing if the ProductID that the user has entered is equal to any of the ProductID from the database in table tblProduct. If it is then an error message will be shown and the message is assigned to the variable validateProduct.
		If cStr(request.form("ProductID")) = cStr(rs("ProductID")) Then
			validateProduct = "This Product ID already exists."
		End if
		rs.movenext
		Loop
		
		LegerCode = request.form("LegerCode")
		LegerCodeExist = False
		
		'Selecting the value in the LegerCode field.
		Select case LegerCode
			Case "73001"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "73002"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "73003"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "73040"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "73060"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "73070"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "73073"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "73090"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "74040"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "79010"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "79020"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "79042"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "79050"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "72031"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
			Case "55090"
				'This if statement is seeing if the variable validateProduct is empty. If it is empty then the system will assign True to the variable LegerCodeExist.
				If validateProduct = "" Then
					LegerCodeExist = True
				End if
		'If the value entered in the LegerCode field is not equal to any of the cases above, then an error message will show.
		Case else
			validateLegerCode = "The Leger code you have entered is not valid. Check the table on the left."
		End select
		'This if statement is seeing if the variable LegerCodeExist is equal to True and if it is then it will go inside the if statement.
		If LegerCodeExist = True Then
			'This if statement is checking if the Code field and the Product VAT field is not left empty so it can then do the following.
			if request.form("Code") <> ""  AND request.form("ProductVAT") <> "" Then
					sql = "INSERT INTO tblProduct (ProductID, LegerCode, SupplierName, Tele, Code, Description, UnitCost, ProductVAT) VALUES ('"&request.form("ProductID")&"', '"&request.form("LegerCode")&"', '"&request.form("SupplierName")&"', '"&request.form("Tele")&"', '"&request.form("Code")&"', '"&request.form("Description")&"', '"&request.form("UnitCost")&"', '"&request.form("ProductVAT")&"')"
					con.execute(sql)
					response.redirect("displayProducts.asp")
			'This if statement is checking if the Code field is not left empty and the Product VAT field is left empty so it can then do the following.
			elseif request.form("Code") <> ""  AND request.form("ProductVAT") = "" Then
					sql = "INSERT INTO tblProduct (ProductID, LegerCode, SupplierName, Tele, Code, Description, UnitCost, ProductVAT) VALUES ('"&request.form("ProductID")&"', '"&request.form("LegerCode")&"', '"&request.form("SupplierName")&"', '"&request.form("Tele")&"', '"&request.form("Code")&"', '"&request.form("Description")&"', '"&request.form("UnitCost")&"', '0')"
					con.execute(sql)
					response.redirect("displayProducts.asp")
			'This if statement is checking if the Code field is left empty and the Product VAT field is not left empty so it can then do the following.
			elseif request.form("Code") = ""  AND request.form("ProductVAT") <> "" Then
					sql = "INSERT INTO tblProduct (ProductID, LegerCode, SupplierName, Tele, Code, Description, UnitCost, ProductVAT) VALUES ('"&request.form("ProductID")&"', '"&request.form("LegerCode")&"', '"&request.form("SupplierName")&"', '"&request.form("Tele")&"', '0', '"&request.form("Description")&"', '"&request.form("UnitCost")&"', '"&request.form("ProductVAT")&"')"
					con.execute(sql)
					response.redirect("displayProducts.asp")
			'If none of the if statements apply, the system will do the following.
			elseif request.form("Code") =""  AND request.form("ProductVAT") = "" Then
					sql = "INSERT INTO tblProduct (ProductID, LegerCode, SupplierName, Tele, Code, Description, UnitCost, ProductVAT) VALUES ('"&request.form("ProductID")&"', '"&request.form("LegerCode")&"', '"&request.form("SupplierName")&"', '"&request.form("Tele")&"', '0', '"&request.form("Description")&"', '"&request.form("UnitCost")&"', '0')"
					con.execute(sql)
					response.redirect("displayProducts.asp")
				end if
		end if
		
	'If the variable pletter is not equal to "P" then this error message will display on the page which is assigned to the variable pletterError.	
	Else
		pletterError = "Product ID must begin with the letter 'P'."
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
    <table width="179" border="1">
      <tr>
        <td width="169"><div class="box">
          <div align="center"><a href="displayTeacherHomepage.asp"><strong>HOMEPAGE</strong></a> </div></td>
        <td width="169"><div class="box">
          <div align="center"><a href="displayProducts.asp"><strong>DISPLAY PRODUCTS</strong></a> </div></td>
      </tr>
    </table>
    <p>----------------------------------------------------------------------------------------------------</p>
  </div>
  <h2 align="center"><u>ADDING A PRODUCT</u></h2>
<table width="200" border="1" align="left">
  <tr>
    <td width="125" height="29"><strong>BUDGET LINE</strong></td>
    <td width="59"><div align="center"><strong>CODE</strong></div></td>
  </tr>
  <tr>
    <td>Stationary</td>
    <td><div align="center"><strong>73001</strong></div></td>
  </tr>
  <tr>
    <td>Printed stationary</td>
    <td><div align="center"><strong>73002</strong></div></td>
  </tr>
  <tr>
    <td>Library books</td>
    <td><div align="center"><strong>73003</strong></div></td>
  </tr>
  <tr>
    <td>Consumables - non teaching</td>
    <td><div align="center"><strong>73040</strong></div></td>
  </tr>
  <tr>
    <td>Hospitality</td>
    <td><div align="center"><strong>73060</strong></div></td>
  </tr>
  <tr>
    <td>Subscription</td>
    <td><div align="center"><strong>73070</strong></div></td>
  </tr>
  <tr>
    <td>Software licences</td>
    <td><div align="center"><strong>73073</strong></div></td>
  </tr>
  <tr>
    <td>Sundry admin expenses</td>
    <td><div align="center"><strong>73090</strong></div></td>
  </tr>
  <tr>
    <td>Travel &amp; Transport</td>
    <td><div align="center"><strong>74040</strong></div></td>
  </tr>
  <tr>
    <td>Consumables</td>
    <td><div align="center"><strong>79010</strong></div></td>
  </tr>
  <tr>
    <td>Text books</td>
    <td><div align="center"><strong>79020</strong></div></td>
  </tr>
  <tr>
    <td>Trips and events</td>
    <td><div align="center"><strong>79042</strong></div></td>
  </tr>
  <tr>
    <td>Photocopying</td>
    <td><div align="center"><strong>79050</strong></div></td>
  </tr>
  <tr>
    <td>Equipment</td>
    <td><div align="center"><strong>72031</strong></div></td>
  </tr>
  <tr>
    <td>Misc. Income</td>
    <td><div align="center"><strong>55090</strong></div></td>
  </tr>
</table>  
<p>&nbsp;</p>
<form name="addProduct" onsubmit="return validateProductForm(this)" method="post" action="addProductForm.asp">
<!--If the error message from the above asp code is needed to display, this will display the error message to the user on the page.-->
	<div><h4><%=validateProduct%></h4></div>
    <div><h4><%=validateLegerCode%></h4></div>
    <div><h4 align="center"><%=pletterError%></h4></div>
  <div align="center">
    <table width="506" border="1">
      <tr>
        <td width="114" height="30"><strong>Product ID</strong></td>
        <td width="368"><input name="ProductID" type="text" id="ProductID" placeholder="Start with P" Required/> 
         *</td>
        </tr>
      <tr>
        <td height="30"><strong>Leger code</strong></td>
        <td> <input name="LegerCode" type="text" id="LegerCode" placeholder="Look at the codes to your left" style="width:300px; height:23px;" > 
        * </td>
      </tr>
      <tr>
        <td height="30"><strong>Supplier name</strong></td>
        <td><input name="SupplierName" type="text" id="SupplierName"> 
        *</td>
        </tr>
      <tr>
        <td height="30"><strong>Tele</strong></td>
        <td><input name="Tele" type="text" id="Tele"> 
         *</td>
        </tr>
      <tr>
        <td height="30"><strong>Code</strong></td>
        <td><input type="text" name="Code" id="Code"></td>
        </tr>
      <tr>
        <td height="30"><strong>Description</strong></td>
        <td><input name="Description" type="text" id="Description" style="width:360px; height:23px;"> 
         *</td>
        </tr>
      <tr>
        <td height="30"><strong>Unit cost</strong></td>
        <td><input name="UnitCost" type="text" id="UnitCost"> 
           *</td>
        </tr>
      <tr>
        <td height="30"><strong>Product VAT</strong></td>
        <td><input name="ProductVAT" type="text" id="ProductVAT"></td>
        </tr>
      <tr>
        <td height="41" colspan="2"><div align="center">
          <p><strong>
            <input name="Submit" type="submit" value="Add"/> 
          </strong></p>
          <p style="color:red"><strong>(*) ARE REQUIRED FIELDS</strong></p>
        </div></td>
        </tr>
    </table>
  </div>
</form>

<p>&nbsp;</p>
</div>

<div id="footer">
Copyright © LSC
</div>
<title></title>
</head>

<body>
</body>
</html>
