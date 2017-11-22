<%
dim Con, rs, sql, OrderID

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

	'This if statement is getting the ProductID1 field from the form and checking if it is not empty. If it is not empty then it will go to the SQL statement. The SQL statement will select everything from the table tblOrder where the TeacherID is the same as the TeacherID that is collected from the log in page and then it will order by OrderID in ascending order.
	if request.form("ProductID1") <> "" Then
		sql = "SELECT * FROM tblOrder WHERE TeacherID = ('"&session("TeacherID")&"') ORDER BY OrderID ASC" 
		rs.Open sql, Con
		
		'This is assigning the OrderID's to the variable OrderID. This will loop until it is end of file. 
		do while not rs.EOF
		OrderID = rs("OrderID")
		rs.movenext
		loop
		
		'This SQL statement is now inserting into the table tblOrderDetail the values that has been entered by the teacher: ProductID1 and Quantity1. The OrderID is also being inserted into the table that's collected from the tblOrder.
		sql = "INSERT INTO tblOrderDetail(ProductID, OrderID, Quantity) VALUES('"&request.form("ProductID1")&"','"&OrderID&"','"&request.form("Quantity1")&"')"
		con.execute(sql)
	end if
	
	'These if statements below will see if the ProductID is not empty and if it isn't then the system will insert into the table tblOrderDetail the values that has been entered by the teacher: ProductID1 and Quantity1. The OrderID is also being inserted into the table that's collected from the tblOrder.
	if request.form("ProductID2") <> "" Then
		sql = "INSERT INTO tblOrderDetail(ProductID, OrderID, Quantity) VALUES('"&request.form("ProductID2")&"','"&OrderID&"','"&request.form("Quantity2")&"')"
		con.execute(sql)
	end if
	
	if request.form("ProductID3") <> "" Then
		sql = "INSERT INTO tblOrderDetail(ProductID, OrderID, Quantity) VALUES('"&request.form("ProductID3")&"','"&OrderID&"','"&request.form("Quantity3")&"')"
		con.execute(sql)
	end if

	if request.form("ProductID4") <> "" Then
		sql = "INSERT INTO tblOrderDetail(ProductID, OrderID, Quantity) VALUES('"&request.form("ProductID4")&"','"&OrderID&"','"&request.form("Quantity4")&"')"
		con.execute(sql)
	end if

	if request.form("ProductID5") <> "" Then
		sql = "INSERT INTO tblOrderDetail(ProductID, OrderID, Quantity) VALUES('"&request.form("ProductID5")&"','"&OrderID&"','"&request.form("Quantity5")&"')"
		con.execute(sql)
	end if

'The system will then redirect the teacher to the submitOrderMessage page where it will confirm their order has been made.
response.redirect("submitOrderMessage.asp")

'Connection is closed.
con.close
%>