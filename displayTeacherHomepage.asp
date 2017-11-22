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

<title>Homepage</title>
</head>

<%
dim Con, rs, sql

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is selecting everything from the table tblOrder where the TeacherID is the same as the TeacherID that is collected from the log in page and then ordering by the OrderID in descending order.
sql = "SELECT * FROM tblOrder WHERE TeacherID = '"& Session("TeacherID")&"' ORDER BY OrderID DESC" 

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
</div>

<div align="right">
<!--This displays the current date.-->  
  <%=Date()%>
</div>

<div id="section">
<div align="center">
  <table width="179" border="1">
    <tr>
      <td width="169">
        <div class="box"> 
        <div align="center"><a href="createOrder.asp"><strong>ADD NEW ORDER</strong></a>
        </div>
        </td>
        <td width="169">
        <div class="box"> 
        <div align="center"><a href="addProductForm.asp"><strong>ADD NEW PRODUCT</strong></a>
        </div>
        </td>
      </tr>
  </table>
  <p>----------------------------------------------------------------------------------------------------</p>
</div>

<table width="500" border="3" align="center" cellpadding="0" cellspacing="0">
    <tr>
    <td height="45" colspan="4"><strong><u>ORDERS PLACED:</u></strong></td>
    </tr>
  <!--This is a loop that will not end until it is end of file. It will select all the OrderID's that the teacher has made by creating an order. Alongside the OrderID, there will be the VIEW, EDIT, DELETE AND THE FINAL ORDER links that will link to specific pages relating to their action.-->
  <%while not rs.EOF%>
  <tr>
    <td width="189"><%=rs("OrderID")%></td>
 	 <td width="101"><div align="center"><A HREF="teacherOrderView.asp?OrderID=<%=rs("OrderID")%>"><font color="black"><strong>VIEW</strong></font></A></div></td>
     <td width="100"><div align="center"><A HREF="editOrder.asp?OrderID=<%=rs("OrderID")%>"><font color="black"><strong>EDIT</strong></font></A></div></td>
     <td width="100"><div align="center"><A HREF="deleteOrderT.asp?OrderID=<%=rs("OrderID")%>"><font color="black"><strong>DELETE</strong></font></A></div></td>  
     <td width="100"><div align="center"><A HREF="OrderViewT.asp?OrderID=<%=rs("OrderID")%>"><font color="black"><strong>FINAL ORDER</strong></font></A></div></td> 
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

</body>
</html>
