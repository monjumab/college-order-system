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
dim Con, rs, sql, TeacherID

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

	sql = "SELECT TeacherID FROM tblOrder WHERE OrderID = " & request.QueryString("OrderID")
	rs.Open sql, Con

	TeacherID = rs("TeacherID")

	rs.close

	sql = "SELECT * FROM tblProduct, tblTeacher, tblOrder, tblOrderDetail WHERE tblTeacher.TeacherID = ('"&TeacherID&"') AND tblProduct.ProductID = tblOrderDetail.ProductID AND tblOrder.OrderID = tblOrderDetail.OrderID AND tblOrderDetail.OrderID = " & request.QueryString("OrderID")
	rs.Open sql, Con
	
	If rs("ChequeNumber") = "0" Then
		response.redirect("notOrderedT.asp")
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
        <div align="center"><a href="displayTeacherHomepage.asp"><strong>HOMEPAGE</strong></a></div>
        </div>
        </td>
      </tr>
  </table>
</div>

<p>----------------------------------------------------------------------------------------------------</p>
<div align="left">
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
        <td width="120"><strong>Cost centre:</strong></td>
        <td width="176"><%=rs("CostCentre")%></td>
      </tr>
      <tr>
        <td width="120"><strong>Section:</strong></td>
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
			dim Product, Total
			Product = ((rs("ProductVAT") + rs("UnitCost"))* rs("Quantity"))
			Total = Total + Product
            response.write(FormatCurrency(Product))
		%>
            </div>
          </td>
        </tr>
		<% rs.movenext
			wend
			rs.MoveFirst
		%>
      </table>
    </div>
    <p>&nbsp;</p>
    <div >
    <table width="312" border="1" align="center">
      <tr>
        <td width="120" align="center" ><strong>Sub-total</strong></td>
        <td width="176" align="center">
          <strong><%=(FormatCurrency(Total))%>
          </strong></td>
      </tr>
      <tr>
        <td align="center"><strong>VAT (20%)</strong></td>
        <td align="center">
          <strong>
          <%
		dim VAT
		VAT = (20 / 100) * Total
		response.write(VAT)
		%>
          </strong></td>
      </tr>
      <tr>
        <td align="center"><strong>Overall total</strong></td>
        <td align="center">
          <strong>
          <%
		dim OverallTotal
		OverallTotal = Total + VAT
		response.write(OverallTotal)
		%>
          </strong></td>
      </tr>
    </table>
    </div>
  </div>
</div>
<div align="center">
      <table width="312" border="1" align="center">
        <tr>
          <td width="120"><div align="center"><strong>Order No.</strong></div></td>
          <td width="176"><div align="center"><strong><%=rs("OrderID")%></strong></div></td>
        </tr>
        <tr>
          <td height="27"><div align="center"><strong>Cheque No.</strong></div></td>
          <td>
          <div align="center">
          <div align='center'><strong>
          <%=rs("ChequeNumber")%>
          </strong></div>
          </div>
          </td>
        </tr>
      </table>
      <p>&nbsp; </p>
    </div>

<div id="footer">
Copyright © LSC
</div>
</body>
</html>
