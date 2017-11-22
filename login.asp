<%
dim Con, rs, sql, Username, Password, GrantAccess, letter, TeacherID

GrantAccess = False

Set Con = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

'Opening a connection to the database.
Con.Open("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=" & Server.MapPath("db/theOrders.accdb"))

'The data that the user has entered in the log in page will be assigned to the variables Username and Password.
Username=request.form("Username")
Password=request.form("Password")

'A successful login is granted false at the moment.
GrantAccess = False

'Whatever the user has entered for the Username, the left most letter will be assigned to the variable letter.
letter = left(Username, 1)

'This if statement is checking if the letter is "t" and if it is then it will select all the tUsername, tPassword and Teacher from the table tblTeacher in the database where the Username the user entered is the same as the tUsername.
if letter = "t" Then
	sql = "SELECT tUsername, tPassword, TeacherID FROM tblTeacher WHERE tUsername = ('"&Username&"')"
	rs.Open sql, Con
	TeacherID = rs("TeacherID")
	Session("TeacherID") = TeacherID
'If the letter is "f" then it will select all the fUsername and fPassword from the table tblFinance in the database where the Username the user entered is the same as the fUsername.
elseif letter = "f" Then
	sql = "SELECT fUsername, fPassword FROM tblFinance WHERE fUsername = ('"&Username&"')"
	rs.Open sql, Con 
End if

'This next if statement is now going to allow the user to successfully log in if their Username and Password is correct. For the  letter "t", if the Username and Password that they entered is the same as the tUsername and tPassword in the database then it will assign True to GrantAccess meaning it will redirect the user to their own teacher homeapage.
if letter = "t" Then
		if Username = rs("tUsername") AND Password = rs("tPassword") Then
			GrantAccess = True
			session("GrantAccess") = Teacher
			response.redirect("displayTeacherHomepage.asp")
		Else
			response.redirect("incorrectLoginMessage.asp")
		End if	
'For the  letter "f", if the Username and Password that they entered is the same as the fUsername and fPassword in the database then it will assign True to GrantAccess meaning it will redirect the user to their Finance homepage.
elseif letter = "f" Then
		if Username = rs("fUsername") AND Password = rs("fPassword") Then
			GrantAccess = True
			session("GrantAccess") = Finance
			response.redirect("displayFinanceHomepage.asp")
		Else
			response.redirect("incorrectLoginMessage.asp")
		End if
'If the letter is not either "t" or "f" then it will redirect to the page where it will display a message saying the log in details are incorrect.
Else
	response.redirect("incorrectLoginMessage.asp")
End if

%>