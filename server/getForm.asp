
<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
set r3=Server.CreateObject("ADODB.Recordset")

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


theConf = trim(request("item"))
if theConf & "a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

sql = "select * from [imknasim].[dbo].[conference] where theIndex=" & theConf
r.open sql,strconn,1,3
if r.eof then
	r.close
	response.write "{}"
else
	s = "{"
	s = s & """theNumber"":""" & j(r("theIndex")) & ""","
	s = s & """theName"":""" & j(r("theName")) & ""","
	s = s & """startdate"":""" & j(r("startdate")) & ""","
	s = s & """enddate"":""" & j(r("enddate")) & ""","
	s = s & """location"":""" & j(r("location")) & ""","
	s = s & """banner"":""" & j(r("banner")) & ""","
	s = s & """logo"":""" & j(r("logo")) & ""","
	s = s & """isourlogo"":""" & b(r("isourlogo")) & ""","
	s = s & """bkcolor"":""" & j(r("bkColor")) & ""","
	s = s & """textcolor"":""" & j(r("textColor")) & ""","
	s = s & """sessionColor"":""" & j(r("sessionColor")) & ""","
	s = s & """splitColor"":""" & j(r("splitColor")) & ""","
	s = s & """textColor2"":""" & j(r("textColor2")) & ""","
	s = s & """bkButton"":""" & j(r("bkButton")) & ""","
	s = s & """logoLight"":""" & b(r("logoLight")) & ""","
	s = s & """roundedCorners"":""" & b(r("roundedcorners")) & ""","
	s = s & """sections"":["
	r.close

	' sections that have fields in this conference.
	' the secondary sort on theindex is required - section.theorder has duplicates
	sql1 = "select distinct s.theindex, s.thename, s.theorder "
	sql1 = sql1 & "from conffield cf, fields f, section s "
	sql1 = sql1 & "where f.theindex=cf.thefield and s.theindex=f.fieldsection "
	sql1 = sql1 & "and cf.theconf=" & theConf & " "
	sql1 = sql1 & "order by s.theorder, s.theindex"
	r1.open sql1,strconn,1,3
	firstSection = true
	while not r1.eof
		if firstSection = false then s = s & ","
		s = s & "{""theNumber"":""" & j(r1("theindex")) & ""","
		s = s & """theName"":""" & j(r1("thename")) & ""","
		s = s & """fields"":["

		' the order is conffield.theorder, not fields.theorder
		sql2 = "select f.theindex, f.thename, f.fieldtype, f.thesize, f.ismandatory, "
		sql2 = sql2 & "f.validation, f.choicetable, cf.theorder as cford "
		sql2 = sql2 & "from conffield cf, fields f "
		sql2 = sql2 & "where f.theindex=cf.thefield and cf.theconf=" & theConf & " "
		sql2 = sql2 & "and f.fieldsection=" & r1("theindex") & " "
		sql2 = sql2 & "order by cford"
		r2.open sql2,strconn,1,3
		firstField = true
		while not r2.eof
			theType = n(r2("fieldtype"))
			if firstField = false then s = s & ","
			s = s & "{""theNumber"":""" & j(r2("theindex")) & ""","
			s = s & """theName"":""" & j(r2("thename")) & ""","
			s = s & """fieldType"":""" & theType & ""","
			s = s & """theSize"":""" & n(r2("thesize")) & ""","
			s = s & """isMandatory"":""" & b(r2("ismandatory")) & ""","
			s = s & """validation"":""" & n(r2("validation")) & ""","
			s = s & """options"":["

			' choice options. field types 7 and 8 ("conference choice")
			' will be added here once a per-conference option table exists
			if theType = "2" or theType = "3" then
				if n(r2("choicetable")) <> "0" then
					sql3 = "select thetable from [imknasim].[dbo].[choicetable] where theindex=" & n(r2("choicetable"))
					r3.open sql3,strconn,1,3
					theTable = ""
					if not r3.eof then theTable = trim(r3("thetable"))
					r3.close
					if theTable <> "" then
						sql3 = "select theindex, thename from " & theTable & " order by thename"
						r3.open sql3,strconn,1,3
						firstOption = true
						while not r3.eof
							if firstOption = false then s = s & ","
							s = s & "{""theNumber"":""" & j(r3("theindex")) & ""","
							s = s & """theName"":""" & j(r3("thename")) & """}"
							firstOption = false
							r3.movenext
						wend
						r3.close
					end if
				end if
			end if

			s = s & "]}"
			firstField = false
			r2.movenext
		wend
		r2.close
		s = s & "]}"
		firstSection = false
		r1.movenext
	wend
	r1.close
	s = s & "]}"
	response.write s
end if
%>
