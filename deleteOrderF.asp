<%
dim Con
Set Con = Server.CreateObject("ADODB.Connection")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is first deleting everything from the table tblOrderDetail where the the OrderID is equal to the OrderID that is requested from the homepage in which Finance would like to delete.
sql = "DELETE * FROM tblOrderDetail WHERE OrderID = " & request.QueryString("OrderID")
con.execute(sql)

'This next SQL statement is then deleting everything from the table tblOrder where the the OrderID is equal to the OrderID that is requested from the homepage in which Finance would like to delete.
sql = "DELETE * FROM tblOrder WHERE OrderID = " & request.QueryString("OrderID")
con.execute(sql)

con.close
'The system will then redirect to the their homepage (Finance).
response.redirect("displayFinanceHomepage.asp")

%>