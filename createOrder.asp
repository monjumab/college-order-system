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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Create Order</title>

</head>

<script type="text/javascript">
//This function is declared so that the field where the date should go in has not been left empty otherwise an error message will pop up alerting the user that they have to enter the date.
function insureDateIsEntered() {
  var OrderDate;
	OrderDate = document.forms["createOrder"]["OrderDate"].value;
	if (OrderDate == '') 
		{
		alert("Order date is required");
		return false;
		}	
}

//This function is used to check if the date is in the correct range.
function validateOrderDateRangeCheck(OrderDate) {
var error = "";
var CurrentDate = new Date();
var datestr = (CurrentDate.getDate()) + "/" + (CurrentDate.getMonth()+1) + "/" + CurrentDate.getFullYear();
  var dayfield=OrderDate.value.split("/")[0]
var monthfield=OrderDate.value.split("/")[1]
var yearfield=OrderDate.value.split("/")[2]
var dayobj = new Date(dayfield, monthfield-1, yearfield)
 
if ((dayfield>31 || dayfield <1 ) || (monthfield > 12 || monthfield < 1)|| (yearfield < 2000) || ((((yearfield % 4)==0)&& monthfield ==2) && dayfield>29) || ((((yearfield % 4)!=0)&& monthfield ==2) && dayfield>28) || (((monthfield % 2) == 0) && dayfield > 30 ) ){
error = "The first Date is invalid.\n";
} else if ((yearfield >= CurrentDate.getFullYear())){
if ( (yearfield = CurrentDate.getFullYear()) ){
if (monthfield >= (CurrentDate.getMonth()+1)) {
if (dayfield > CurrentDate.getDate()){
OrderDate.style.background = '#1F00FF'; 
error = "The first date is greater than the current date.\n";
}
} 
}
}
return error;   
}
 
//This function is used to check that the format for the date that the user has entered is correct.
function validateOrderDate(OrderDate) {
var error = "";
 
if (/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.test(OrderDate.value)==false|| /[^0-9]+$/.test(OrderDate.value)  ) {
error = "The date is in the wrong format(DD/MM/YYYY).\n";
}
return error;   
}
  </script>

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
  <%=Date()%>
</div>

<div align="center">
  <table width="122" border="1">
    <tr>
      <td width="169">
        <div class="box"> 
        <div align="center"><a href="displayTeacherHomepage.asp"><strong>HOMEPAGE</strong></a>
        </div>
        </td>
      </tr>
  </table>
  <p>----------------------------------------------------------------------------------------------------</p></div>

<div>
	<form name="createOrder" method="post" onsubmit="return insureDateIsEntered()" action="addOrder.asp"> 
    	 <p align="center"><strong>Insert Date:</strong>
  		 <input type="text" name="OrderDate" value="<%=Date()%>" placeholder="dd/mm/yyyy"/>
  		 <div align="center">
   <p>
     <input type="submit" name="Submit" value="Add">
      </p>
   <p>&nbsp;</p>
  </div>
    </form>
</div>

<div id="footer">
Copyright © LSC
</div>

</body>
</html>
