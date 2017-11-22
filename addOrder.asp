<%
dim Con

Set Con = Server.CreateObject("ADODB.Connection")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

	'This SQL statement is creating an order by inserting the TeacherID, the date they had entered and the ChequeNumber to '0'. It will then redirect to the addForm page.
	sql = "INSERT INTO tblOrder(TeacherID, OrderDate, ChequeNumber) VALUES('"&Session("TeacherID")&"', '"& request.form("OrderDate")&"', '0')"
	con.execute(sql)
	
response.redirect("addForm.asp")	
	
%>