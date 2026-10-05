
<!-- #include file="secure.inc" -->
<!-- #include file="lib.inc" -->
<%
' live duplicate check while the registrant is still filling the form.
' answers one question only - is this mobile already registered for this
' conference - and deliberately returns nothing else about the registration.
set r=Server.CreateObject("ADODB.Recordset")

theConf = trim(request("item") & "")
if theConf & "a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

theCell = normCell(request("cell"))

if theConf = "0" or theCell = "" then
	response.write "{""exists"":""False""}"
else
	sql = "select theindex from [imknasim].[dbo].[registration] "
	sql = sql & "where theconf=" & theConf & " and cellnum='" & replace(theCell, "'", "''") & "'"
	r.open sql,strconn,1,3
	if r.eof then
		response.write "{""exists"":""False""}"
	else
		response.write "{""exists"":""True""}"
	end if
	r.close
end if
%>
