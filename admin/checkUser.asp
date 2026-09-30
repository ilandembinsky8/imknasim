<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")

login=request.form("username")
password=request.form("password")

sql = "select * from [imknasim].[dbo].[users] where login='"&login&"' and password='"&password&"'"
r.open sql,strconn,1,3
if not r.eof then
    session("user")=login
    response.redirect "inner.asp"
else
    session("user")=""
    response.redirect "index.asp"
end if
%>