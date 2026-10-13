
<!-- #include file="secure.inc" -->
<!-- #include file="lib.inc" -->
<%
' Conference-specific choice fields (field types 7 and 8).
' There is no mapping in the data - fields.choicetable is 0 on all of them and
' choicetable holds no row for the new per-conference tables - so the source of
' each one is named here. Keep them together so the mapping stays visible.
const FLD_DAYS          = 28
const FLD_SESSIONS      = 29
const FLD_DISCOUNT      = 33
const FLD_COLLECTPOINT  = 48
const FLD_COMPANIONDAYS = 68
' field 30 "special activities" has no source yet and is deliberately not rendered

set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
set r3=Server.CreateObject("ADODB.Recordset")

dim secId(200)
dim secName(200)
dim secGrp(200)
dim secCs(200)

' two sections are merged when they share the same conferencesection.thenumber
' and that number is not 0
function sameGroup(a, other)
	if a = other then
		sameGroup = true
	elseif secGrp(a) <> "0" and secGrp(a) = secGrp(other) then
		sameGroup = true
	else
		sameGroup = false
	end if
end function

' a merged section is emitted once, at the member holding the lowest
' conferencesection.theindex. that member also gives it its name
function isGroupLeader(idx, total)
	dim k
	isGroupLeader = true
	for k = 0 to total - 1
		if k <> idx then
			if sameGroup(idx, k) then
				if secCs(k) < secCs(idx) then isGroupLeader = false
			end if
		end if
	next
end function

' types 7 and 8 are rendered only for the fields whose source is known
function isKnownConfField(fid)
	isKnownConfField = false
	if fid = cstr(FLD_DAYS) then isKnownConfField = true
	if fid = cstr(FLD_SESSIONS) then isKnownConfField = true
	if fid = cstr(FLD_DISCOUNT) then isKnownConfField = true
	if fid = cstr(FLD_COLLECTPOINT) then isKnownConfField = true
	if fid = cstr(FLD_COMPANIONDAYS) then isKnownConfField = true
end function

' append option rows that already carry theindex and thename
sub emitSimple(sqlText)
	dim firstOne
	r3.open sqlText,strconn,1,3
	firstOne = true
	while not r3.eof
		if firstOne = false then s = s & ","
		s = s & "{""theNumber"":""" & j(r3("theindex")) & ""","
		s = s & """theName"":""" & j(r3("thename")) & """}"
		firstOne = false
		r3.movenext
	wend
	r3.close
end sub

' the generic lookup used by ordinary choice fields, driven by choicetable
sub emitChoiceTable(choiceId)
	dim t
	if n(choiceId) = "0" then exit sub
	r3.open "select thetable from [imknasim].[dbo].[choicetable] where theindex=" & n(choiceId),strconn,1,3
	t = ""
	if not r3.eof then t = trim(r3("thetable") & "")
	r3.close
	if t = "" then exit sub
	call emitSimple("select theindex, thename from " & t & " order by thename")
end sub

' one option per day of the conference. the value is the date so it stays
' sortable and unambiguous, the label is what the registrant reads
sub emitDays(fromText, toText)
	dim d1, d2, d, firstOne
	d1 = parseDbDate(fromText)
	d2 = parseDbDate(toText)
	if isempty(d1) then exit sub
	if isempty(d2) then d2 = d1
	if d2 < d1 then d2 = d1
	firstOne = true
	d = d1
	do while d <= d2
		if firstOne = false then s = s & ","
		s = s & "{""theNumber"":""" & dbDate(d) & ""","
		s = s & """theName"":""" & j(dayLabel(d)) & """}"
		firstOne = false
		d = dateadd("d", 1, d)
	loop
end sub

theConf = trim(request("item"))
if theConf & "a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

sql = "select * from [imknasim].[dbo].[conference] where theIndex=" & theConf
r.open sql,strconn,1,3
if r.eof then
	r.close
	response.write "{}"
else
	confStart = trim(r("startdate") & "")
	confEnd   = trim(r("enddate") & "")
	confEdit  = b(r("isedit"))
	confClose = n(r("whentoclose"))
	editOpen  = editIsOpen(r("isedit"), confStart, confClose)

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
	r.close

	' editing after registration is offered only when the conference allows it
	s = s & """isEdit"":""" & confEdit & ""","
	if confEdit = "True" then
		theDeadline = editDeadline(confStart, confClose)
		if isempty(theDeadline) then
			s = s & """editUntil"":"""",""editUntilRaw"":"""","
		else
			s = s & """editUntil"":""" & showDate(theDeadline) & ""","
			s = s & """editUntilRaw"":""" & dbDate(theDeadline) & ""","
		end if
	else
		s = s & """editUntil"":"""",""editUntilRaw"":"""","
	end if
	if editOpen then
		s = s & """editOpen"":""True"","
	else
		s = s & """editOpen"":""False"","
	end if
	s = s & """sections"":["

	' sections holding fields in this conference, together with their merge
	' group from conferencesection. the secondary sort on theindex is required -
	' section.theorder has duplicate values
	sql1 = "select distinct s.theindex, s.thename, s.theorder, "
	sql1 = sql1 & "isnull(cs.thenumber, 0) as grp, isnull(cs.theindex, 999999) as csidx "
	sql1 = sql1 & "from conffield cf "
	sql1 = sql1 & "inner join fields f on f.theindex = cf.thefield "
	sql1 = sql1 & "inner join section s on s.theindex = f.fieldsection "
	sql1 = sql1 & "left join conferencesection cs on cs.thesection = s.theindex "
	sql1 = sql1 & "and cs.theconference = " & theConf & " "
	sql1 = sql1 & "where cf.theconf = " & theConf & " "
	sql1 = sql1 & "order by s.theorder, s.theindex"
	r1.open sql1,strconn,1,3
	cntSec = 0
	while not r1.eof
		secId(cntSec)   = n(r1("theindex"))
		secName(cntSec) = trim(r1("thename") & "")
		secGrp(cntSec)  = n(r1("grp"))
		secCs(cntSec)   = clng(n(r1("csidx")))
		cntSec = cntSec + 1
		r1.movenext
	wend
	r1.close

	firstSection = true
	for i = 0 to cntSec - 1
		if isGroupLeader(i, cntSec) then

			idList = ""
			for k = 0 to cntSec - 1
				if sameGroup(i, k) then
					if idList <> "" then idList = idList & ","
					idList = idList & secId(k)
				end if
			next

			if firstSection = false then s = s & ","
			s = s & "{""theNumber"":""" & secId(i) & ""","
			s = s & """theName"":""" & j(secName(i)) & ""","
			s = s & """fields"":["

			' conffield.theorder is a single order across the whole conference,
			' so merged sections interleave correctly with no extra work
			sql2 = "select f.theindex, f.thename, f.fieldtype, f.thesize, f.ismandatory, "
			sql2 = sql2 & "f.validation, f.choicetable, cf.theorder as cford "
			sql2 = sql2 & "from conffield cf, fields f "
			sql2 = sql2 & "where f.theindex=cf.thefield and cf.theconf=" & theConf & " "
			sql2 = sql2 & "and f.fieldsection in (" & idList & ") "
			sql2 = sql2 & "order by cford"
			r2.open sql2,strconn,1,3
			firstField = true
			while not r2.eof
				theType = n(r2("fieldtype"))
				theFieldId = n(r2("theindex"))

				renderIt = true
				if theType = "7" or theType = "8" then
					if isKnownConfField(theFieldId) = false then renderIt = false
				end if

				if renderIt then
					if firstField = false then s = s & ","
					s = s & "{""theNumber"":""" & theFieldId & ""","
					s = s & """theName"":""" & j(r2("thename")) & ""","
					s = s & """fieldType"":""" & theType & ""","
					s = s & """theSize"":""" & n(r2("thesize")) & ""","
					s = s & """isMandatory"":""" & b(r2("ismandatory")) & ""","
					s = s & """validation"":""" & n(r2("validation")) & ""","
					s = s & """options"":["

					if theType = "2" or theType = "3" then
						call emitChoiceTable(r2("choicetable"))
					elseif theType = "7" or theType = "8" then
						if theFieldId = cstr(FLD_DISCOUNT) then
							call emitSimple("select theindex, thename from [imknasim].[dbo].[discount] where theconference=" & theConf & " order by theindex")
						elseif theFieldId = cstr(FLD_COLLECTPOINT) then
							call emitSimple("select theindex, thename from [imknasim].[dbo].[collectpoint] where theconference=" & theConf & " order by theindex")
						elseif theFieldId = cstr(FLD_SESSIONS) then
							call emitSimple("select theindex, thename from [imknasim].[dbo].[session] where conference=" & theConf & " and showinagenda=1 order by theday, theorder")
						elseif theFieldId = cstr(FLD_DAYS) or theFieldId = cstr(FLD_COMPANIONDAYS) then
							call emitDays(confStart, confEnd)
						end if
					end if

					s = s & "]}"
					firstField = false
				end if

				r2.movenext
			wend
			r2.close
			s = s & "]}"
			firstSection = false
		end if
	next

	s = s & "]}"
	response.write s
end if
%>
