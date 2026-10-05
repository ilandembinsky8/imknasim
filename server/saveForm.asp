
<!-- #include file="secure.inc" -->
<!-- #include file="lib.inc" -->
<%
' saves a registration, or updates an existing one when reg + code are posted.
' values go through the recordset, never into SQL text.
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
set r3=Server.CreateObject("ADODB.Recordset")
set r4=Server.CreateObject("ADODB.Recordset")

' the mobile field is the identity key. field 9 (landline) is never used for
' identity - it is often a shared institutional number
const CELL_FIELD  = 11
const EMAIL_FIELD = 8
const FIRST_FIELD = 1
const LAST_FIELD  = 2

dim fId(300)
dim fType(300)
dim fMand(300)
dim fChoice(300)
dim vals(50)
dim labels(50)

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

' write the answers of one field in place: reuse existing rows, add what is
' missing, delete only the surplus. a failure here can never empty the whole
' registration the way a delete-then-reinsert would
sub writeField(regId, fieldId, cntVals)
	dim sqlw, i
	sqlw = "select * from [imknasim].[dbo].[regfield] "
	sqlw = sqlw & "where thereg=" & regId & " and thefield=" & fieldId & " order by theindex"
	r2.open sqlw,strconn,1,3
	i = 0
	while not r2.eof
		if i < cntVals then
			r2("thevalue") = left(vals(i), 255)
			if labels(i) <> "" then
				r2("thetext") = left(labels(i), 255)
			else
				r2("thetext") = null
			end if
			r2.update
			i = i + 1
		else
			r2.delete
		end if
		r2.movenext
	wend
	while i < cntVals
		r2.addnew
		r2("thereg") = clng(regId)
		r2("thefield") = clng(fieldId)
		r2("thevalue") = left(vals(i), 255)
		if labels(i) <> "" then r2("thetext") = left(labels(i), 255)
		r2.update
		i = i + 1
	wend
	r2.close
end sub

theConf = trim(request.form("item") & "")
if theConf & "a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

theReg = trim(request.form("reg") & "")
if not isnumeric(theReg) then theReg = ""
theCode = trim(request.form("code") & "")

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

	theCell = normCell(posted(CELL_FIELD))

	if missing <> "" then
		missing = left(missing, len(missing) - 1)
		response.write "{""ok"":""False"",""error"":""missing"",""fields"":""" & missing & """}"
	elseif theCell = "" then
		response.write "{""ok"":""False"",""error"":""missing"",""fields"":""" & CELL_FIELD & """}"
	else
		' editing requires the code of that registration, otherwise anyone could
		' post a registration number and overwrite someone else's answers
		authOk = true
		isEdit = false
		if theReg <> "" then
			isEdit = true
			sql = "select * from [imknasim].[dbo].[registration] "
			sql = sql & "where theindex=" & theReg & " and theconf=" & theConf
			r.open sql,strconn,1,3
			if r.eof then
				authOk = false
			elseif trim(r("thecode") & "") = "" or trim(r("thecode") & "") <> theCode then
				authOk = false
			end if
			r.close
		end if

		if authOk = false then
			response.write "{""ok"":""False"",""error"":""badcode""}"
		else
			' the mobile must not already belong to another registration of this conference
			sql = "select theindex from [imknasim].[dbo].[registration] "
			sql = sql & "where theconf=" & theConf & " and cellnum='" & replace(theCell, "'", "''") & "'"
			if isEdit = true then sql = sql & " and theindex<>" & theReg
			r.open sql,strconn,1,3
			taken = not r.eof
			r.close

			if taken = true then
				response.write "{""ok"":""False"",""error"":""duplicate""}"
			else
				theFull = trim(posted(FIRST_FIELD) & " " & posted(LAST_FIELD))
				theMail = posted(EMAIL_FIELD)

				if isEdit = true then
					sql = "select * from [imknasim].[dbo].[registration] where theindex=" & theReg
					r.open sql,strconn,1,3
					r("thename") = left(theFull, 255)
					r("cellnum") = left(theCell, 50)
					r("theemail") = left(theMail, 50)
					r.update
					outCode = trim(r("thecode") & "")
					r.close
				else
					outCode = newCode()
					r.open "registration",strconn,1,3
					r.addnew
					r("theconf") = clng(theConf)
					r("thename") = left(theFull, 255)
					r("cellnum") = left(theCell, 50)
					r("theemail") = left(theMail, 50)
					r("thedate") = left(stampNow(), 10)
					r("thetime") = right(stampNow(), 5)
					r("thecode") = outCode
					r("thetries") = 0
					r.update
					theReg = n(r("theindex"))
					r.close

					if theReg = "0" then
						sql = "select max(theindex) as newid from [imknasim].[dbo].[registration] where theconf=" & theConf
						r.open sql,strconn,1,3
						if not r.eof then theReg = n(r("newid"))
						r.close
					end if
				end if

				for i = 0 to cnt - 1
					theVal = posted(fId(i))
					cntVals = 0
					if theVal <> "" then
						if fType(i) = "3" then
							' multi choice arrives comma separated, one row per selection
							parts = split(theVal, ",")
							for k = 0 to ubound(parts)
								onePart = trim(parts(k))
								if onePart <> "" then
									vals(cntVals) = onePart
									labels(cntVals) = choiceText(fChoice(i), onePart)
									cntVals = cntVals + 1
								end if
							next
						elseif fType(i) = "2" then
							if theVal <> "0" then
								vals(0) = theVal
								labels(0) = choiceText(fChoice(i), theVal)
								cntVals = 1
							end if
						else
							vals(0) = theVal
							labels(0) = ""
							cntVals = 1
						end if
					end if
					call writeField(theReg, fId(i), cntVals)
				next

				s = "{""ok"":""True"",""theNumber"":""" & theReg & ""","
				if isEdit = true then
					s = s & """mode"":""edit""}"
				else
					s = s & """mode"":""new"",""theCode"":""" & outCode & """}"
				end if
				response.write s
			end if
		end if
	end if
end if
%>
