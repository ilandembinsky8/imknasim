
<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
set r3=Server.CreateObject("ADODB.Recordset")
set r4=Server.CreateObject("ADODB.Recordset")

dim fId(300)
dim fType(300)
dim fMand(300)
dim fChoice(300)

' escape text for JSON
function j(v)
	dim sj
	if isnull(v) then
		sj = ""
	else
		sj = trim(cstr(v))
	end if
	sj = replace(sj, "\", "\\")
	sj = replace(sj, """", "\""")
	sj = replace(sj, vbCrLf, " ")
	sj = replace(sj, vbCr, " ")
	sj = replace(sj, vbLf, " ")
	sj = replace(sj, vbTab, " ")
	j = sj
end function

' bit to True/False, same convention as the rest of the project
function b(v)
	if isnull(v) then
		b = "False"
	elseif v = true then
		b = "True"
	else
		b = "False"
	end if
end function

' number, empty becomes 0
function n(v)
	if isnull(v) then
		n = "0"
	else
		n = trim(cstr(v))
	end if
end function

' posted value of one field, always a trimmed string
function posted(fieldId)
	posted = trim(request.form("f" & fieldId) & "")
end function

' resolve the label of a choice value on the server rather than trusting the
' browser, so regfield.thetext records what was really offered
function choiceText(choiceTableId, theValue)
	dim t, sqlc
	choiceText = ""
	if n(choiceTableId) = "0" then exit function
	if not isnumeric(theValue) then exit function
	sqlc = "select thetable from [imknasim].[dbo].[choicetable] where theindex=" & n(choiceTableId)
	r4.open sqlc,strconn,1,3
	t = ""
	if not r4.eof then t = trim(r4("thetable") & "")
	r4.close
	if t = "" then exit function
	sqlc = "select thename from " & t & " where theindex=" & theValue
	r4.open sqlc,strconn,1,3
	if not r4.eof then choiceText = trim(r4("thename") & "")
	r4.close
end function

' one regfield row. values go through the recordset, never into SQL text
sub saveValue(regId, fieldId, theValue, theLabel)
	r2.open "regfield",strconn,1,3
	r2.addnew
	r2("thereg") = clng(regId)
	r2("thefield") = clng(fieldId)
	r2("thevalue") = left(theValue, 255)
	if theLabel <> "" then r2("thetext") = left(theLabel, 255)
	r2.update
	r2.close
end sub

theConf = trim(request.form("item") & "")
if theConf & "a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

sql = "select theIndex from [imknasim].[dbo].[conference] where theIndex=" & theConf
r.open sql,strconn,1,3
confFound = not r.eof
r.close

if confFound = false then
	response.write "{""ok"":""False"",""error"":""conference not found""}"
else
	' load the fields this conference actually has. anything posted that is not
	' on this list is ignored, so a crafted POST cannot write rows for other fields
	cnt = 0
	sql1 = "select f.theindex, f.fieldtype, f.ismandatory, f.choicetable "
	sql1 = sql1 & "from conffield cf, fields f "
	sql1 = sql1 & "where f.theindex=cf.thefield and cf.theconf=" & theConf
	r1.open sql1,strconn,1,3
	while not r1.eof
		fId(cnt) = n(r1("theindex"))
		fType(cnt) = n(r1("fieldtype"))
		fMand(cnt) = b(r1("ismandatory"))
		fChoice(cnt) = n(r1("choicetable"))
		cnt = cnt + 1
		r1.movenext
	wend
	r1.close

	' mandatory check. client side validation can be bypassed
	missing = ""
	for i = 0 to cnt - 1
		if fMand(i) = "True" then
			theVal = posted(fId(i))
			if fType(i) = "2" and theVal = "0" then theVal = ""
			if theVal = "" then missing = missing & fId(i) & ","
		end if
	next

	if missing <> "" then
		missing = left(missing, len(missing) - 1)
		response.write "{""ok"":""False"",""error"":""missing"",""fields"":""" & missing & """}"
	else
		' registration.thename / cellnum / theemail come from fixed superset field ids:
		' 1 first name, 2 last name, 8 email, 11 mobile, 9 phone as a fallback
		theFull = trim(posted(1) & " " & posted(2))
		theMail = posted(8)
		thePhone = posted(11)
		if thePhone = "" then thePhone = posted(9)

		r.open "registration",strconn,1,3
		r.addnew
		r("theconf") = clng(theConf)
		r("thename") = left(theFull, 255)
		r("cellnum") = left(thePhone, 50)
		r("theemail") = left(theMail, 50)
		r("thedate") = year(date) & "-" & right("0" & month(date), 2) & "-" & right("0" & day(date), 2)
		r("thetime") = right("0" & hour(time), 2) & ":" & right("0" & minute(time), 2)
		r.update
		theReg = n(r("theindex"))
		r.close

		' fallback in case the provider does not hand back the identity on update
		if theReg = "0" then
			sql = "select max(theindex) as newid from [imknasim].[dbo].[registration] where theconf=" & theConf
			r.open sql,strconn,1,3
			if not r.eof then theReg = n(r("newid"))
			r.close
		end if

		for i = 0 to cnt - 1
			theVal = posted(fId(i))
			if theVal <> "" then
				if fType(i) = "3" then
					' multi choice arrives comma separated, one row per selection
					parts = split(theVal, ",")
					for k = 0 to ubound(parts)
						onePart = trim(parts(k))
						if onePart <> "" then
							call saveValue(theReg, fId(i), onePart, choiceText(fChoice(i), onePart))
						end if
					next
				elseif fType(i) = "2" then
					if theVal <> "0" then
						call saveValue(theReg, fId(i), theVal, choiceText(fChoice(i), theVal))
					end if
				else
					call saveValue(theReg, fId(i), theVal, "")
				end if
			end if
		next

		response.write "{""ok"":""True"",""theNumber"":""" & theReg & """}"
	end if
end if
%>
