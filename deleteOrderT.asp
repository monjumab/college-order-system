<%
dim Con
Set Con = Server.CreateObject("ADODB.Connection")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'This SQL statement is first deleting everything from the table tblOrderDetail where the the OrderID is equal to the OrderID that is requested from the homepage in which the teacher would like to delete.
sql1 = "DELETE * FROM tblOrderDetail WHERE OrderID = " & request.QueryString("OrderID")
con.execute(sql1)

'This next SQL statement is then deleting everything from the table tblOrder where the the OrderID is equal to the OrderID that is requested from the homepage in which the teacher would like to delete.
sql2 = "DELETE * FROM tblOrder WHERE OrderID = " & request.QueryString("OrderID")
con.execute(sql2)

'The system will then redirect to the their homepage (Teacher).
response.redirect("displayTeacherHomepage.asp")

con.close
%>