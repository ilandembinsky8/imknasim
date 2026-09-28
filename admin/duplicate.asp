<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
	indexField = ""
	subTable = ""

thetable = request("page")
currpage = thetable
allFields = request("all")

%>
<!-- #include file="menu.inc" -->
<%
sql = "SELECT * FROM "&thetable&" where "&indexField&"="&request("num")
r.Open sql,strconn,1,3
r1.open thetable,strconn,2,3
r1.addnew
for i=1 to numOfFields
	r1(eval("field"&i)) = r(eval("field"&i))
next
r1.update
r.close
response.redirect "inner.asp?page="&request("page")
%>
	