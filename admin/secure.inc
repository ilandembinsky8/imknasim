<%
On Error Resume Next
strConn = "DSN=imknasim;Uid=michal;Pwd=Im$$2026;"
'strConn1 = "DSN=old_stamps;Uid=stamps;Pwd=aaa$$$333;"

'set myconn = Server.CreateObject ("ADODB.Connection")
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset") 


'If Request.ServerVariables("SERVER_PORT")=80 Then
'    Dim strSecureURL
'    strSecureURL = "https://"
'    strSecureURL = strSecureURL & Request.ServerVariables("SERVER_NAME")
'    strSecureURL = strSecureURL & Request.ServerVariables("URL")
'    Response.Redirect strSecureURL
'End If
%>

