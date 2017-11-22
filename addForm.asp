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
    padding:50px;	 	 
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
#submit {
 color: #585858;
 margin: auto;
 font-size: 40;
 width: 75px;
 height: 45px;
 padding: 0px;
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
    height: 50px;
    border: 3px solid black;
	text-align:center;
	color: black;
	margin: auto;
	font-size:20px
}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Add order</title>

<script type="text/javascript">
//This function is declared in order to validate the ProductID and Quantity fields if they are left empty and to ensure that the Quantity is a number between 1 and 50.
function validateOrderForm() {
	//These variables are declared by getting the data from the form labelled ID.
  var ProductID, Quantity;
	ProductID1 = document.forms["addOrder"]["ProductID1"].value;
	Quantity1 = document.forms["addOrder"]["Quantity1"].value;
	ProductID2 = document.forms["addOrder"]["ProductID2"].value;
	Quantity2 = document.forms["addOrder"]["Quantity2"].value;
	ProductID3 = document.forms["addOrder"]["ProductID3"].value;
	Quantity3 = document.forms["addOrder"]["Quantity3"].value;
	ProductID4 = document.forms["addOrder"]["ProductID4"].value;
	Quantity4 = document.forms["addOrder"]["Quantity4"].value;
	ProductID5 = document.forms["addOrder"]["ProductID5"].value;
	Quantity5 = document.forms["addOrder"]["Quantity5"].value;
	
	//This if statement is going to alert the user with an error message if both the ProductID1 and Quantity1 fields are left empty.
	if (ProductID1 == "" && Quantity1 == "") 
		{
		alert( "ProductID1 and Quantity1 is required.");
		return false;
		}
	//This if statement is going to alert the user with an error message if the ProductID1 field is left empty.
	if (ProductID1 == "") 
		{
		alert("ProductID1 is required.");
		return false;
		}
	//This if statement is going to alert the user with an error message if the Quantity1 field is left empty.	
	if (Quantity1 == "") 
		{
		alert("Quantity1 is required.");
		return false;
		}	
		
	//These if statements below are to check if both the ProductID and Quantity fields are filled in if one of them is already filled in and the other is left empty. 
	if (ProductID2 != "" && Quantity2 == "")
		{
		alert("Please fill in the Quantity2 field.")	
		return false;
		}
	if (Quantity2 != "" && ProductID2 == "")
		{
		alert("Please fill in the ProductID2 field.")	
		}
		
	if (ProductID3 != "" && Quantity3 == "")
		{
		alert("Please fill in the Quantity3 field.")	
		return false;
		}
	if (Quantity3 != "" && ProductID3 == "")
		{
		alert("Please fill in the ProductID3 field.")	
		}
		
	if (ProductID4 != "" && Quantity4 == "")
		{
		alert("Please fill in the Quantity4 field.")	
		return false;
		}
	if (Quantity4 != "" && ProductID4 == "")
		{
		alert("Please fill in the ProductID4 field.")	
		}	
		
	if (ProductID5 != "" && Quantity5 == "")
		{
		alert("Please fill in the Quantity5 field.")	
		return false;
		}
	if (Quantity5 != "" && ProductID5 == "")
		{
		alert("Please fill in the ProductID5 field.")	
		}	
	
	
	//These if statements below are to check if the value entered in the Quantity fields is a number and if it is then it will check whether that number is between 1 and 50.
	if (/^\d+(\.\d{0})?$/.test(Quantity1)==false)
		{
		alert("Quantity1 must be a number.");
		return false
		}
		if (Quantity1 <1 || Quantity1 > 50)
		{
		alert("Quantity1 must contain a number between 1 and 50")
		return false	
		}
	if (ProductID2 != "" && Quantity2 != "")
		{
			if (/^\d+(\.\d{0})?$/.test(Quantity2)==false)
			{
			alert("Quantity2 must be a number.");
			return false
			}	
			if (Quantity2 <1 || Quantity2 > 50)
			{
			alert("Quantity2 must contain a number between 1 and 50")
			return false	
			}
		}
	if (ProductID3 != "" && Quantity3 != "")
		{
			if (/^\d+(\.\d{0})?$/.test(Quantity3)==false)
			{
			alert("Quantity3 must be a number.");
			return false
			}	
			if (Quantity3 <1 || Quantity3 > 50)
			{
			alert("Quantity3 must contain a number between 1 and 50")
			return false	
			}
		}
	if (ProductID4 != "" && Quantity4 != "")
		{
			if (/^\d+(\.\d{0})?$/.test(Quantity4)==false)
			{
			alert("Quantity4 must be a number.");
			return false
			}
			if (Quantity4 <1 || Quantity4 > 50)
			{
			alert("Quantity4 must contain a number between 1 and 50")
			return false	
			}
		}
	if (ProductID5 != "" && Quantity5 != "")
		{
			if (/^\d+(\.\d{0})?$/.test(Quantity5)==false)
			{
			alert("Quantity5 must be a number.");
			return false
			}
			if (Quantity5 <1 || Quantity5 > 50)
			{
			alert("Quantity5 must contain a number between 1 and 50")
			return false	
			}	
		}			
}
		
</script>

<%
dim Con, rs, sql

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is selecting everything from the table tblTeacher and tblProduct.
sql = "SELECT * FROM tblTeacher, tblProduct WHERE TeacherID = ('"&session("TeacherID")&"')"

rs.Open sql, Con

%>

</head>

<body>

<div id="header">
<div align="center">
  <table width="419" align="center">
    <tr>
      <td width="127" height="60" align="center"><strong><a href="logout.asp">LOG OUT</a></strong></td>
      <th width="280" rowspan="5">
      <img src="http://www.jobs.ac.uk/images/employer-logos/medium/1164.gif" alt="l" width="152" height="124" align="left" /></th>
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

<h2 align="center"><strong><u>CREATE ORDER</u></strong></h2>

<div align="center">
      <table width="134" border="1">
        <tr>
          <td width="124">
            <div class="box"> 
            <div align="center"><a href="displayProducts.asp"><strong>VIEW PRODUCTS</strong></a>
            </div>
          </td>
        </tr>
  </table>
    </div>
  <p align="center">----------------------------------------------------------------------------------------------------</p>
  <table width="382" align="center">
    <tr>
  <td width="165">
      <div align="center"><strong>First name:</strong> <%=rs("FirstName")%>
    </div></td>
  </tr>
  <tr>
    <td>
      <div align="center"><strong>Last name:</strong> <%=rs("LastName")%>
    </div></td>
  </tr>
  <tr>
    <td>
      <div align="center"><strong>Delivery room:</strong> <%=rs("DeliveryRoom")%>
    </div></td>
  </tr>
  <tr>
    <td>
      <div align="center"><strong>Cost centre:</strong> <%=rs("CostCentre")%>
    </div></td>
  </tr>
  <tr>
    <td>
      <div align="center"><strong>Section:</strong> <%=rs("Section")%>
    </div></td>
  </tr>
</table>

  <form name="addOrder" onsubmit="return validateOrderForm(this)" method="post" action="addOrderDetail.asp?TeacherID=<%=rs("TeacherID")%>">
  <p align="center">&nbsp;</p> 
  <p align="center"><strong>Product ID1:</strong>
    <input name="ProductID1" type="text" id="ProductID1"> 
    <strong>Quantity1: </strong>
    <input name="Quantity1" type="text" id="Quantity1">
    <br>
    <br>
    <strong>Product ID2:</strong>
    <input type="text" name="ProductID2">
    <strong>Quantity2: </strong>
    <input name="Quantity2" type="text" id="Quantity2">
    <br>
    <br>
    <strong>Product ID3:</strong>
    <input type="text" name="ProductID3">
    <strong>Quantity3: </strong>
    <input name="Quantity3" type="text" id="Quantity3">
    <br>
    <br>
    <strong>Product ID4:</strong>
    <input type="text" name="ProductID4">
    <strong>Quantity4: </strong>
    <input name="Quantity4" type="text" id="Quantity4">
    <br>
    <br>
    <strong>Product ID5:</strong>
    <input type="text" name="ProductID5">
    <strong>Quantity5: </strong>
    <input name="Quantity5" type="text" id="Quantity5">
  </p>
  <div align="center">
   <p>
     <input name="Submit" type="submit" value="Add">
      </p>
   <p>&nbsp;</p>
  </div>
</form>


<div id="footer">
Copyright © LSC
</div>
</body>
</html>
