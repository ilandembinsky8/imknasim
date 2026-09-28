<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
sql = "delete from "&request("table")&" where theindex="&request("item")
r.open sql,strconn,1,3
response.redirect "forumDetails.asp?num="&request("num")&"&item="&request("table")
%>
	