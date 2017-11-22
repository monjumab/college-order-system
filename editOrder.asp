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
    height: 45px;
    border: 3px solid black;
	text-align:center;
	color: black;
	margin: auto;
	font-size:20px
}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Edit order</title>
</head>

<%
dim Con, rs, sql

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is selecting everything from tables tblProduct, tblTeacher, tblOrder and tblOrderDetail where the TeacherID in tblTeacher is the same as the TeacherID from the login page. The rest of the statement is linking the tables using a primary key in one table and a foreign key in another.
sql = "SELECT * FROM tblProduct, tblTeacher, tblOrder, tblOrderDetail WHERE tblTeacher.TeacherID = ('"&session("TeacherID")&"') AND tblProduct.ProductID = tblOrderDetail.ProductID AND tblOrder.OrderID = tblOrderDetail.OrderID AND tblOrderDetail.OrderID = " & request.QueryString("OrderID")

rs.Open sql, Con

'This if statement is checking if the ChequeNumber is not equal to 0. If it isn't then it will redirect the page to the page where it will say it is already ordered.
If rs("ChequeNumber") <> "0" Then
		response.redirect("alreadyOrderedT.asp")
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
  <table width="109" border="1">
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

<h2 align="center"><strong><u>EDIT ORDER</u></strong></h2>
<div id="section">
<!--This is a loop that will not end until it is end of file.It will display all the Product ID's and its quantity that the teacher has ordered. A link to the EDIT page will be next to that product if the teacher wishes to edit that specific product.-->
  <%while not rs.EOF%>
  <div align="center">
    <table width="430" height="32" border="1">
      <tr>
        <td width="105"><strong>Product ID:</strong></td>
        <td width="56"><%=rs("ProductID")%></td>
        <td width="73"><strong>Quantity:</strong></td>
        <td width="116"><%=rs("Quantity")%></td>
        <td width="46"><strong><a href="updateOrder.asp?OrderDetailID=<%=rs("OrderDetailID")%>">EDIT</a></strong></td>
      </tr>
    </table>
  </div>
  <%rs.movenext
  wend
  %>
  <p>&nbsp;</p>
</div>

<div id="footer">
Copyright © LSC
</div>
</body>
</html>
