
<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
set r3=Server.CreateObject("ADODB.Recordset")

' בריחה של טקסט לתוך JSON
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

' bit אל True/False, כמו שאר הפרויקט
function b(v)
	if isnull(v) then
		b = "False"
	elseif v = true then
		b = "True"
	else
		b = "False"
	end if
end function

' מספר, כשריק מוחזר 0
function n(v)
	if isnull(v) then
		n = "0"
	else
		n = trim(cstr(v))
	end if
end function

theConf = trim(request("item"))
if theConf&"a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

sql = "select * from [imknasim].[dbo].[conference] where theIndex="&theConf
r.open sql,strconn,1,3
if r.eof then
	r.close
	response.write "{}"
else
	s = "{""theNumber"":"""&j(r("theIndex"))&""",""theName"":"""&j(r("theName"))&_
	    """,""startdate"":"""&j(r("startdate"))&""",""enddate"":"""&j(r("enddate"))&_
	    """,""location"":"""&j(r("location"))&""",""banner"":"""&j(r("banner"))&_
	    """,""logo"":"""&j(r("logo"))&""",""isourlogo"":"""&b(r("isourlogo"))&_
	    """,""bkcolor"":"""&j(r("bkColor"))&""",""textcolor"":"""&j(r("textColor"))&_
	    """,""sessionColor"":"""&j(r("sessionColor"))&""",""splitColor"":"""&j(r("splitColor"))&_
	    """,""textColor2"":"""&j(r("textColor2"))&""",""bkButton"":"""&j(r("bkButton"))&_
	    """,""logoLight"":"""&b(r("logoLight"))&""",""roundedCorners"":"""&b(r("roundedcorners"))&_
	    """,""sections"":["
	r.close

	' סקשנים שיש להם שדות בכנס הזה. המיון המשני לפי theindex הכרחי -
	' בטבלת section יש ערכי theorder כפולים
	sql1 = "select distinct s.theindex, s.thename, s.theorder from conffield cf, fields f, section s "&_
	       "where f.theindex=cf.thefield and s.theindex=f.fieldsection and cf.theconf="&theConf&" "&_
	       "order by s.theorder, s.theindex"
	r1.open sql1,strconn,1,3
	firstSection = true
	while not r1.eof
		if firstSection = false then s = s & ","
		s = s & "{""theNumber"":"""&j(r1("theindex"))&""",""theName"":"""&j(r1("thename"))&""",""fields"":["

		' הסדר הוא conffield.theorder ולא fields.theorder
		sql2 = "select f.theindex, f.thename, f.fieldtype, f.thesize, f.ismandatory, "&_
		       "f.validation, f.choicetable, cf.theorder as cford "&_
		       "from conffield cf, fields f "&_
		       "where f.theindex=cf.thefield and cf.theconf="&theConf&" "&_
		       "and f.fieldsection="&r1("theindex")&" order by cford"
		r2.open sql2,strconn,1,3
		firstField = true
		while not r2.eof
			theType = n(r2("fieldtype"))
			if firstField = false then s = s & ","
			s = s & "{""theNumber"":"""&j(r2("theindex"))&""",""theName"":"""&j(r2("thename"))&_
			    """,""fieldType"":"""&theType&""",""theSize"":"""&n(r2("thesize"))&_
			    """,""isMandatory"":"""&b(r2("ismandatory"))&""",""validation"":"""&n(r2("validation"))&_
			    """,""options"":["

			' אפשרויות בחירה. סוגים 7 ו-8 ("בחירה - כנס") יתווספו כאן
			' כשתיווצר טבלת רשימות ספציפיות לכנס
			if theType = "2" or theType = "3" then
				if n(r2("choicetable")) <> "0" then
					sql3 = "select thetable from [imknasim].[dbo].[choicetable] where theindex="&n(r2("choicetable"))
					r3.open sql3,strconn,1,3
					theTable = ""
					if not r3.eof then theTable = trim(r3("thetable"))
					r3.close
					if theTable <> "" then
						sql3 = "select theindex, thename from "&theTable&" order by thename"
						r3.open sql3,strconn,1,3
						firstOption = true
						while not r3.eof
							if firstOption = false then s = s & ","
							s = s & "{""theNumber"":"""&j(r3("theindex"))&""",""theName"":"""&j(r3("thename"))&"""}"
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
