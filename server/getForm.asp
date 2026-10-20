
<!-- #include file="secure.inc" -->
<!-- #include file="lib.inc" -->
<%
' Conference-specific choice fields (types 7 and 8).
' Fields backed by a real table are wired through choicetable like any other
' choice field - the only difference is that their table is filtered by
' theconference. The two below are computed instead and have no table, so they
' stay named here.
const FLD_DAYS          = 28
const FLD_SESSIONS      = 29
const FLD_COMPANIONDAYS = 68
' the yes/no field that carries a roommate name beside it
const FLD_PARTNER       = 43
' field 30 "special activities" has no source yet and is not rendered

' the exact option label that reveals the free text companion field
const OTHER_LABEL = "אחר"

set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
set r3=Server.CreateObject("ADODB.Recordset")
set r4=Server.CreateObject("ADODB.Recordset")

dim secId(200)
dim secName(200)
dim secGrp(200)
dim secCs(200)

' theindex of the option literally named "אחר" in the field just emitted,
' or "0" when that field has no such option
otherId = "0"

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

' the two computed conference choices, which have no backing table
function isComputedConfField(fid)
	isComputedConfField = false
	if fid = cstr(FLD_DAYS) then isComputedConfField = true
	if fid = cstr(FLD_SESSIONS) then isComputedConfField = true
	if fid = cstr(FLD_COMPANIONDAYS) then isComputedConfField = true
end function

' append option rows that carry theindex and thename
sub emitSimple(sqlText)
	dim firstOne
	r3.open sqlText,strconn,1,3
	firstOne = true
	while not r3.eof
		if firstOne = false then s = s & ","
		s = s & "{""theNumber"":""" & j(r3("theindex")) & ""","
		s = s & """theName"":""" & j(r3("thename")) & """}"
		if trim(r3("thename") & "") = OTHER_LABEL then otherId = j(r3("theindex"))
		firstOne = false
		r3.movenext
	wend
	r3.close
end sub

' options from a per-conference table. confhotel also carries a price and a
' stock count, and a room with nothing left is not offered at all
sub emitConfTable(t)
	dim sqlc, firstOne, isHotel
	isHotel = (lcase(t) = "confhotel")
	sqlc = "select * from [imknasim].[dbo].[" & t & "] where theconference=" & theConf
	if isHotel then sqlc = sqlc & " and leftquantity > 0"
	sqlc = sqlc & " order by theindex"
	r3.open sqlc,strconn,1,3
	firstOne = true
	while not r3.eof
		if firstOne = false then s = s & ","
		s = s & "{""theNumber"":""" & j(r3("theindex")) & ""","
		s = s & """theName"":""" & j(r3("thename")) & """"
		if isHotel then
			' thePrice and breakfast are per night. withPartner is a flag: it
			' says the room may be shared, and sharing does not change the rate
			s = s & ",""thePrice"":""" & n(r3("theprice")) & """"
			s = s & ",""withPartner"":""" & b(r3("withpartner")) & """"
			s = s & ",""breakfast"":""" & n(r3("breakfast")) & """"
			s = s & ",""leftQuantity"":""" & n(r3("leftquantity")) & """"
		end if
		s = s & "}"
		if trim(r3("thename") & "") = OTHER_LABEL then otherId = j(r3("theindex"))
		firstOne = false
		r3.movenext
	wend
	r3.close
end sub

' the generic lookup, driven by choicetable. a per-conference table is filtered
' by theconference, a shared lookup table is not
sub emitChoiceTable(choiceId)
	dim t
	if n(choiceId) = "0" then exit sub
	r4.open "select thetable from [imknasim].[dbo].[choicetable] where theindex=" & n(choiceId),strconn,1,3
	t = ""
	if not r4.eof then t = trim(r4("thetable") & "")
	r4.close
	if t = "" then exit sub
	if isConfScoped(t) then
		call emitConfTable(t)
	else
		call emitSimple("select theindex, thename from " & t & " order by thename")
	end if
end sub

' tables that hold one set of rows per conference rather than a shared list
function isConfScoped(t)
	dim lt
	lt = lcase(trim(t))
	isConfScoped = false
	if lt = "discount" then isConfScoped = true
	if lt = "collectpoint" then isConfScoped = true
	if lt = "confhotel" then isConfScoped = true
end function

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

	' closeform shuts the form by hand, whenclose shuts it on a date
	confShut = "False"
	if formIsShut(r("closeform"), r("whenclose")) then confShut = "True"

	' which of the three confprice columns today falls into
	confTier = priceTier(confStart, r("whenearly"))

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

	' a closed form is still described, so the page can show the conference
	' heading above the notice instead of an empty screen
	s = s & """closeForm"":""" & confShut & ""","
	s = s & """priceTier"":""" & confTier & ""","

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

	' registration price per participant type, so the form can fill the price
	' field as soon as a participant type is chosen. the three columns are
	' resolved to one figure here rather than in the browser, so that the
	' visitor's clock cannot move the price and saveForm cannot disagree
	s = s & """prices"":["
	sql = "select theparticipant, theprice, earlyprice, confdayprice "
	sql = sql & "from [imknasim].[dbo].[confprice] "
	sql = sql & "where theconference=" & theConf & " order by theparticipant"
	r4.open sql,strconn,1,3
	firstPrice = true
	while not r4.eof
		if firstPrice = false then s = s & ","
		s = s & "{""theParticipant"":""" & n(r4("theparticipant")) & ""","
		s = s & """thePrice"":""" & pickPrice(confTier, r4("earlyprice"), r4("theprice"), r4("confdayprice")) & """}"
		firstPrice = false
		r4.movenext
	wend
	r4.close
	s = s & "],"

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
			' so merged sections interleave correctly with no extra work.
			' conffield.ismandatory lets one conference add a mandatory field
			' without changing the shared definition
			sql2 = "select f.theindex, f.thename, f.fieldtype, f.thesize, f.ismandatory, "
			sql2 = sql2 & "f.validation, f.choicetable, cf.theorder as cford, "
			sql2 = sql2 & "isnull(cf.ismandatory, 0) as confmand "
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
					if n(r2("choicetable")) = "0" then
						if isComputedConfField(theFieldId) = false then renderIt = false
					end if
				end if

				if renderIt then
					' mandatory at either level makes the field mandatory
					theMand = "False"
					if b(r2("ismandatory")) = "True" then theMand = "True"
					if b(r2("confmand")) = "True" then theMand = "True"

					otherId = "0"

					if firstField = false then s = s & ","
					s = s & "{""theNumber"":""" & theFieldId & ""","
					s = s & """theName"":""" & j(r2("thename")) & ""","
					s = s & """fieldType"":""" & theType & ""","
					s = s & """theSize"":""" & n(r2("thesize")) & ""","
					s = s & """isMandatory"":""" & theMand & ""","
					s = s & """validation"":""" & n(r2("validation")) & ""","
					s = s & """options"":["

					if theType = "2" or theType = "3" then
						call emitChoiceTable(r2("choicetable"))
					elseif theType = "7" or theType = "8" then
						if n(r2("choicetable")) <> "0" then
							call emitChoiceTable(r2("choicetable"))
						elseif theFieldId = cstr(FLD_SESSIONS) then
							call emitSimple("select theindex, thename from [imknasim].[dbo].[session] where conference=" & theConf & " and showinagenda=1 order by theday, theorder")
						elseif theFieldId = cstr(FLD_DAYS) or theFieldId = cstr(FLD_COMPANIONDAYS) then
							call emitDays(confStart, confEnd)
						end if
					end if

					s = s & "],"
					' the roommate name is the same kind of companion, but its owner
					' is a yes/no field, so the value that opens it is True. whether
					' the chosen room may be shared at all is decided in the browser,
					' from withPartner, and enforced again by saveForm
					if theFieldId = cstr(FLD_PARTNER) then otherId = "True"
					' a single choice field offering "אחר" gets a free text
					' companion, posted and stored as 10000 + this field id
					' a multi choice field is deliberately left out: its owner is a
					' group of checkboxes with no single answer to key a companion on
					if theType = "2" or theType = "7" or theFieldId = cstr(FLD_PARTNER) then
						s = s & """otherOption"":""" & otherId & """}"
					else
						s = s & """otherOption"":""0""}"
					end if
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
