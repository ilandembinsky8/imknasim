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
    if r("userType")=1 then 
        session("userType")="admin"
    else
        session("userType")="user"
    end if
    response.redirect "inner.asp"
    'response.write r("thename")
else
    session("user")=""
    response.redirect "index.asp"
    'response.write "failure"
end if
%>