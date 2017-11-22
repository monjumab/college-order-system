<%
'This is abondoning the whole log in and then redirecting the user to the log in page.
session.abandon()
response.redirect("adminPage.html")
%>