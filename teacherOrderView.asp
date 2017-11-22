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
<title>Order view</title>

<%
dim Con, rs, sql

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is selecting everything from the table tblProduct, tblTeacher, tblOrder and tblOrderDetail where the TeacherID from the table tblTeacher is equal to the variable TeacherID. The statement is then linking all the tables together using a primary key in one table and a foreign key in another.
sql = "SELECT * FROM tblProduct, tblTeacher, tblOrder, tblOrderDetail WHERE tblTeacher.TeacherID = ('"&session("TeacherID")&"') AND tblProduct.ProductID = tblOrderDetail.ProductID AND tblOrder.OrderID = tblOrderDetail.OrderID AND tblOrderDetail.OrderID = " & request.QueryString("OrderID")

rs.Open sql, Con

'This if statement is seeing if the ChequeNumber from the databse is not equal to 0. If it is not equal to 0 then the system will redirect the user to the page where it will display a message saying the order has already been ordered.
If rs("ChequeNumber") <> "0" Then
		response.redirect("alreadyOrderedT.asp")
End if

%>

</head>

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
        <div align="center"><a href="displayTeacherHomepage.asp"><strong>HOMEPAGE</strong></a>
        </div>
        </td>
      </tr>
  </table>
 <p>----------------------------------------------------------------------------------------------------</p> 
</div>

<h2> ORDER NUMBER 
  <div align="center">
  <table width="200" border="3">
  <tr>
  	<!--This is displaying the OrderID of that order.-->
    <td><div align="center"><%=rs("OrderID")%></div></td>
  </tr>
</table>
</div>
</h2>
  <div align="left">
  	<!--The following is displaying the details of the teacher.-->
    <table width="312" border="3" align="right">
      <tr>
        <td width="120"><strong>Ordered by:</strong></td>
        <td width="176"><%=rs("FirstName")%> <%=rs("LastName")%></td>
      </tr>
      <tr>
        <td width="120"><strong>Date: </strong></td>
        <td width="176"><%=rs("OrderDate")%></td>
      </tr>
      <tr>
        <td width="120"><strong>Delivery room:</strong></td>
        <td width="176"><%=rs("DeliveryRoom")%></td>
      </tr>
    </table>
    <table width="312" border="1">
      <tr>
        <td width="120"><strong>Cost centre: </strong></td>
        <td width="176"><%=rs("CostCentre")%></td>
      </tr>
      <tr>
        <td width="120"><strong>Section: </strong></td>
        <td width="176"><%=rs("Section")%></td>
      </tr>
    </table>
    <p>&nbsp;</p>
    <div align="center">
      <table width="1243" border="1">
        <tr>
          <td width="85"><div align="center"><strong>Leger code</strong></div></td>
          <td width="125"><div align="center"><strong>Supplier name</strong></div></td>
          <td width="130"><div align="center"><strong>Tele</strong></div></td>
          <td width="210"><div align="center"><strong>Code</strong></div></td>
          <td width="270"><div align="center"><strong>Description</strong></div></td>
          <td width="105"><div align="center"><strong>Product VAT</strong></div></td>
          <td width="85"><div align="center"><strong>Unit cost</strong></div></td>
          <td width="85"><div align="center"><strong>Quantity</strong></div></td>
          <td width="90"><div align="center"><strong>Total</strong></div></td>
        </tr>
        <tr>
          <!--This is a loop that will not end until end of file. This will display all the products and their details from the table tblProduct from the database that's related to what the teacher has ordered.. It will return back the ProductID, LegerCode, SupplierName, Tele, Code, Description, UnitCost and ProductVAT.-->
          <%while not rs.eof%>
          <td><div align="center"><%=rs("LegerCode")%></div></td>
          <td><div align="center"><%=rs("SupplierName")%></div></td>
          <td><div align="center"><%=rs("Tele")%></div></td>
          <td><div align="center"><%=rs("Code")%></div></td>
          <td><div align="center"><%=rs("Description")%></div></td>
          <td><div align="center"><%=rs("ProductVAT")%></div></td>
          <td><div align="center"><%=rs("UnitCost")%></div></td>
          <td><div align="center"><%=rs("Quantity")%></div></td>
          <td>
            <div align="center">
        <%
			'The following is calculating the cost of the product by adding together the Product VAT with the Unit Cost and then multiplying the answer by the amount the teacher wanted and then assigning it to the variable Product.			
			dim Product, Total
			Product = ((rs("ProductVAT") + rs("UnitCost"))* rs("Quantity"))
			'This is adding all the values in the variable Product together and then assigning it to the variable Total.
			Total = Total + Product
			'This is formatting the value in the variable Product to currency.            
			response.write(FormatCurrency(Product))
		%>
            </div>
          </td>
        </tr>
		<% rs.movenext
			wend
		%>
      </table>
    </div>
    <p>&nbsp;</p>
    <div >
    <table width="312" border="1" align="center">
      <tr>
        <td width="120" align="center" ><strong>Sub-total</strong></td>
        <td width="176" align="center">
          <!--This is formatting the value in the variable Total to currency.-->        
          <strong><%=(FormatCurrency(Total))%>
          </strong></td>
      </tr>
      <tr>
        <td align="center"><strong>VAT (20%)</strong></td>
        <td align="center">
          <strong>
        <%
		'The following is working out the VAT cost. VAT is 20%. 20 is divided by 100 and then multiplying the answer by the value in variable Total.
		dim VAT
		VAT = (20 / 100) * Total
		'This is then formatting the value in the variable VAT to currency.		
		response.write(FormatCurrency(VAT))
		%>
          </strong></td>
      </tr>
      <tr>
        <td align="center"><strong>Overall total</strong></td>
        <td align="center">
          <strong>
        <%
		'The following is working out the overall cost of the order. It's done by adding together the value in the variable Total and the value in the variable VAT. The answer is then assigned to the variable OverallTotal.
		dim OverallTotal
		OverallTotal = Total + VAT
		'This is then formatting the value in the variable OverallTotal.
		response.write(FormatCurrency(OverallTotal))
		%>
          </strong></td>
      </tr>
    </table>
    </div>
    <p>&nbsp;</p>
    <p>&nbsp;</p> 
  </div>

</div>

<div id="footer">
Copyright © LSC
</div>
</body>
</html>
