
<!-- #include file="secure.inc" -->
<!-- #include file="lib.inc" -->
<%
' sign in to an existing registration with the mobile number and the 4 digit
' code, and hand back the saved answers so the form can be prefilled.
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")

theConf = trim(request("item") & "")
if theConf & "a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

theCell = normCell(request("cell"))
theCode = trim(request("code") & "")

if theConf = "0" or theCell = "" or theCode = "" then
	response.write "{""ok"":""False"",""error"":""input""}"
else
	sql = "select * from [imknasim].[dbo].[registration] "
	sql = sql & "where theconf=" & theConf & " and cellnum='" & replace(theCell, "'", "''") & "'"
	r.open sql,strconn,1,3

	if r.eof then
		r.close
		' same answer as a wrong code, so this cannot be used to discover
		' which numbers are registered
		response.write "{""ok"":""False"",""error"":""badcode""}"
	else
		theReg = n(r("theindex"))
		savedCode = trim(r("thecode") & "")
		tries = n(r("thetries"))
		lockUntil = trim(r("theuntil") & "")

		if lockUntil <> "" and lockUntil > stampNow() then
			r.close
			response.write "{""ok"":""False"",""error"":""locked"",""until"":""" & j(lockUntil) & """}"
		elseif savedCode = "" or savedCode <> theCode then
			' wrong code - count it, and lock the registration once the limit is hit
			tries = cint(tries) + 1
			if tries >= MAX_TRIES then
				r("thetries") = 0
				r("theuntil") = stampAt(dateadd("n", LOCK_MINUTES, now))
			else
				r("thetries") = tries
			end if
			r.update
			r.close
			response.write "{""ok"":""False"",""error"":""badcode""}"
		else
			' correct code - clear the counter and return the saved answers
			r("thetries") = 0
			r("theuntil") = ""
			r.update
			r.close

			s = "{""ok"":""True"",""theNumber"":""" & theReg & """,""values"":["
			sql = "select thefield, thevalue from [imknasim].[dbo].[regfield] "
			sql = sql & "where thereg=" & theReg & " order by thefield, theindex"
			r1.open sql,strconn,1,3
			lastField = ""
			acc = ""
			firstOut = true
			while not r1.eof
				thisField = n(r1("thefield"))
				thisValue = trim(r1("thevalue") & "")
				if thisField <> lastField then
					if lastField <> "" then
						if firstOut = false then s = s & ","
						s = s & "{""theField"":""" & lastField & ""","
						s = s & """theValue"":""" & j(acc) & """}"
						firstOut = false
					end if
					lastField = thisField
					acc = thisValue
				else
					' a multi choice field has one row per selection
					acc = acc & "," & thisValue
				end if
				r1.movenext
			wend
			if lastField <> "" then
				if firstOut = false then s = s & ","
				s = s & "{""theField"":""" & lastField & ""","
				s = s & """theValue"":""" & j(acc) & """}"
			end if
			r1.close
			s = s & "]}"
			response.write s
		end if
	end if
end if
%>
