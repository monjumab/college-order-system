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
    width: 170px;
    height: 50px;
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

<%
dim Con, rs, sql

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is selecting everything from the table tblProduct.
sql = "SELECT * FROM tblProduct"

rs.Open sql, Con

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
<td>&nbsp;</td>
<div align="center"></div>
</p>
</div>

<div id="section">
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
        <div align="center"><a href="addProductForm.asp"><strong>ADD NEW PRODUCTS</strong></a>
        </div>
        </td>
      </tr>
  </table>
  <p>----------------------------------------------------------------------------------------------------</p>
</div>

<p>&nbsp;</p>
<table width="1000" border="3" align="center">
  <tr>
    <td width="76"><strong>Product ID</strong></td>
    <td width="55"><strong>Leger Code</strong></td>
    <td width="102"><strong>Supplier Name</strong></td>
    <td width="100"><strong>Tele</strong></td>
    <td width="117"><strong>Code</strong></td>
    <td width="218"><strong>Description</strong></td>
    <td width="94"><strong>Unit Cost</strong></td>
    <td width="94"><strong>Product VAT</strong></td>
  </tr>
  <!--This is a loop that will not end until it is end of file. This will display all the products and their details from the table tblProduct from the database. It will return back the ProductID, LegerCode, SupplierName, Tele, Code, Description, UnitCost and ProductVAT.-->
  <% while not rs.EOF%>
  <tr>
    <td><strong><%=rs("ProductID")%></strong></td>
    <td><%=rs("LegerCode")%></td>
    <td><%=rs("SupplierName")%></td>
    <td><%=rs("Tele")%></td>
    <td><%=rs("Code")%></td>
    <td><%=rs("Description")%></td>
    <td><%=rs("UnitCost")%></td>
    <td><%=rs("ProductVAT")%></td>
  </tr>
  <%
  rs.movenext
  wend
  %>
</table>

</div>

<div id="footer">
Copyright © LSC
</div>
<title></title>
</head>

<body>
</body>
</html>
