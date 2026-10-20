
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
const FLD_DAYS          = 28
const FLD_SESSIONS      = 29
const FLD_DISCOUNT      = 33
const FLD_COLLECTPOINT  = 48
const FLD_COMPANIONDAYS = 68

' the registration price is the conference price plus the stay, and every
' figure it is built from lives in confprice or confhotel
const FLD_PARTICIPANT   = 25
const FLD_PRICE         = 32
const FLD_NEEDLODGING   = 38
const FLD_CHECKIN       = 39
const FLD_CHECKOUT      = 40
const FLD_NIGHTS        = 41
const FLD_ROOM          = 42
const FLD_PARTNER       = 43
const FLD_BREAKFAST     = 44

' the exact option label that requires a free text companion field
const OTHER_LABEL = "אחר"
' companion answers are stored as this offset plus the field id
const OTHER_OFFSET = 10000
' the form's own bookkeeping rides in regfield on two further bands, so
' that no column has to be added to registration. both are keyed on the
' room field: STOCK says this registration consumed a unit of confhotel
' stock, PAIR names the registration it shares that room with. real field
' ids never reach 20000, so these can never collide with an answer
const STOCK_OFFSET = 20000
const PAIR_OFFSET  = 21000

const CELL_FIELD  = 11
const EMAIL_FIELD = 8
const FIRST_FIELD = 1
const LAST_FIELD  = 2

dim fId(300)
dim fType(300)
dim fMand(300)
dim fChoice(300)
dim fOther(300)
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

' theindex of the option literally named "אחר" for this choice field,
' or "0" when the field offers no such option
function otherOptionId(choiceTableId)
	dim t, sqlc
	otherOptionId = "0"
	if n(choiceTableId) = "0" then exit function
	sqlc = "select thetable from [imknasim].[dbo].[choicetable] where theindex=" & n(choiceTableId)
	r4.open sqlc,strconn,1,3
	t = ""
	if not r4.eof then t = trim(r4("thetable") & "")
	r4.close
	if t = "" then exit function
	sqlc = "select theindex from " & t & " where thename=N'" & OTHER_LABEL & "'"
	r4.open sqlc,strconn,1,3
	if not r4.eof then otherOptionId = n(r4("theindex"))
	r4.close
end function

' the table a choice field reads from, lowercased, or an empty string
function choiceTableName(choiceTableId)
	dim sqlc
	choiceTableName = ""
	if n(choiceTableId) = "0" then exit function
	sqlc = "select thetable from [imknasim].[dbo].[choicetable] where theindex=" & n(choiceTableId)
	r4.open sqlc,strconn,1,3
	if not r4.eof then choiceTableName = lcase(trim(r4("thetable") & ""))
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

' label for a conference-specific choice value (field types 7 and 8).
' resolved on the server, never taken from the browser
function confChoiceText(fieldId, theValue)
	dim sqlc, d
	confChoiceText = ""
	if fieldId = cstr(FLD_DAYS) or fieldId = cstr(FLD_COMPANIONDAYS) then
		d = parseDbDate(theValue)
		if isempty(d) then exit function
		confChoiceText = dayLabel(d)
		exit function
	end if
	if not isnumeric(theValue) then exit function
	if fieldId = cstr(FLD_DISCOUNT) then
		sqlc = "select thename from [imknasim].[dbo].[discount] where theindex=" & theValue
	elseif fieldId = cstr(FLD_COLLECTPOINT) then
		sqlc = "select thename from [imknasim].[dbo].[collectpoint] where theindex=" & theValue
	elseif fieldId = cstr(FLD_SESSIONS) then
		sqlc = "select thename from [imknasim].[dbo].[session] where theindex=" & theValue
	else
		exit function
	end if
	r4.open sqlc,strconn,1,3
	if not r4.eof then confChoiceText = trim(r4("thename") & "")
	r4.close
end function

' a yes/no field answers True or False and nothing else. a hand built post
' saying "true" or "1" reads as False here, so it can neither slip past the
' checks that look for the exact answer nor be stored as a third value
' the posted answer of field number i, normalised for its type
function answerOf(i)
	if fType(i) = "4" then
		answerOf = postedFlag(fId(i))
	else
		answerOf = posted(fId(i))
	end if
end function

function postedFlag(fieldId)
	postedFlag = "False"
	if posted(fieldId) = "True" then postedFlag = "True"
end function

' trim, and collapse runs of spaces, so that a roommate typed as
' "ploni  almoni " still matches a registration stored as "ploni almoni"
function squash(v)
	dim t
	t = cstr(v & "")
	t = replace(t, vbTab, " ")
	t = replace(t, vbCr, " ")
	t = replace(t, vbLf, " ")
	t = trim(t)
	do while instr(t, "  ") > 0
		t = replace(t, "  ", " ")
	loop
	squash = t
end function

' a roommate can only be declared when lodging is wanted and the chosen
' room is one the client marked as shareable. everything about sharing
' hangs off this, on the server as well as in the browser
function roomSharable()
	dim roomId, sqlr
	roomSharable = false
	if postedFlag(FLD_NEEDLODGING) <> "True" then exit function
	roomId = posted(FLD_ROOM)
	if roomId = "" then exit function
	if roomId = "0" then exit function
	if not isnumeric(roomId) then exit function
	sqlr = "select withpartner from [imknasim].[dbo].[confhotel] "
	sqlr = sqlr & "where theconference=" & theConf & " and theindex=" & clng(roomId)
	r2.open sqlr,strconn,1,3
	if not r2.eof then
		if b(r2("withpartner")) = "True" then roomSharable = true
	end if
	r2.close
end function

' the other registration of this conference that is on the same room and
' either carries the name typed here, or typed this registration's name
' itself. one that is already paired with somebody else is skipped, so a
' room never ends up holding three people
function findPartner(myReg, roomId, partnerName, myName)
	dim sqlm, nm, mine, roomCol, meId
	findPartner = 0
	if roomId = "" then exit function
	if not isnumeric(roomId) then exit function
	nm = squash(partnerName)
	if nm = "" then exit function
	mine = squash(myName)
	roomCol = hotelField
	if roomCol = "0" then exit function
	' a registration being created has no number yet, and there is then no
	' row of its own to exclude
	meId = 0
	if isnumeric(myReg) then meId = clng(myReg)

	sqlm = "select top 1 rg.theindex from [imknasim].[dbo].[registration] rg "
	sqlm = sqlm & "where rg.theconf=" & theConf & " and rg.theindex<>" & meId & " "
	sqlm = sqlm & "and not exists (select 1 from [imknasim].[dbo].[regfield] pr "
	sqlm = sqlm & "  where pr.thereg=rg.theindex and pr.thefield=" & (PAIR_OFFSET + clng(roomCol))
	sqlm = sqlm & "  and rtrim(pr.thevalue)<>'' and rtrim(pr.thevalue)<>'" & meId & "') "
	sqlm = sqlm & "and exists (select 1 from [imknasim].[dbo].[regfield] f "
	sqlm = sqlm & "  where f.thereg=rg.theindex and f.thefield=" & roomCol
	sqlm = sqlm & "  and rtrim(f.thevalue)='" & clng(roomId) & "') "
	sqlm = sqlm & "and (ltrim(rtrim(rg.thename))=N'" & replace(nm, "'", "''") & "' "
	sqlm = sqlm & "  or exists (select 1 from [imknasim].[dbo].[regfield] g "
	sqlm = sqlm & "    where g.thereg=rg.theindex and g.thefield=" & (OTHER_OFFSET + FLD_PARTNER)
	sqlm = sqlm & "    and ltrim(rtrim(g.thevalue))=N'" & replace(mine, "'", "''") & "')) "
	sqlm = sqlm & "order by rg.theindex"
	r2.open sqlm,strconn,1,3
	if not r2.eof then findPartner = clng(n(r2("theindex")))
	r2.close
end function

' the regfield id that carries a note for the room field in use
function noteId(band)
	noteId = 0
	if hotelField = "0" then exit function
	noteId = band + clng(hotelField)
end function

' read one bookkeeping note, or "" when there is none
function noteGet(regId, fieldId)
	dim sqlg
	noteGet = ""
	if fieldId = 0 then exit function
	if clng(n(regId)) = 0 then exit function
	sqlg = "select thevalue from [imknasim].[dbo].[regfield] "
	sqlg = sqlg & "where thereg=" & clng(n(regId)) & " and thefield=" & fieldId
	r3.open sqlg,strconn,1,3
	if not r3.eof then noteGet = trim(r3("thevalue") & "")
	r3.close
end function

' write one bookkeeping note. an empty value removes the row, so that a
' registration holding nothing carries no note at all
sub noteSet(regId, fieldId, theNote)
	dim sqls
	if fieldId = 0 then exit sub
	if clng(n(regId)) = 0 then exit sub
	sqls = "select * from [imknasim].[dbo].[regfield] "
	sqls = sqls & "where thereg=" & clng(n(regId)) & " and thefield=" & fieldId
	r3.open sqls,strconn,1,3
	if r3.eof then
		if theNote <> "" then
			r3.addnew
			r3("thereg") = clng(n(regId))
			r3("thefield") = fieldId
			r3("thevalue") = theNote
			r3.update
		end if
	else
		if theNote = "" then
			r3.delete
		else
			r3("thevalue") = theNote
			r3.update
		end if
	end if
	r3.close
end sub

' is that registration still stored as being on this room
function roomStillWanted(regId, roomId)
	dim sqlw
	roomStillWanted = false
	if hotelField = "0" then exit function
	if not isnumeric(roomId) then exit function
	sqlw = "select thevalue from [imknasim].[dbo].[regfield] "
	sqlw = sqlw & "where thereg=" & clng(n(regId)) & " and thefield=" & hotelField
	r3.open sqlw,strconn,1,3
	if not r3.eof then
		if trim(r3("thevalue") & "") = cstr(roomId) then roomStillWanted = true
	end if
	r3.close
end function

' record whether a registration is consuming a unit of confhotel stock
sub setHolding(regId, flag)
	dim note
	note = ""
	if flag then note = "True"
	call noteSet(regId, noteId(STOCK_OFFSET), note)
end sub

' point a registration at the one it shares a room with, or at nobody
sub setPartner(regId, partnerId)
	dim note
	note = ""
	if clng(n(partnerId)) <> 0 then note = cstr(clng(n(partnerId)))
	call noteSet(regId, noteId(PAIR_OFFSET), note)
end sub

' drop a link only if it still points back here, so that changing partners
' cannot blank out a pairing somebody else has since made
sub clearPartner(regId, meReg)
	if noteGet(regId, noteId(PAIR_OFFSET)) = cstr(clng(n(meReg))) then
		call noteSet(regId, noteId(PAIR_OFFSET), "")
	end if
end sub

' does that registration currently hold the room it is on
function partnerHoldsRoom(partnerReg, roomId)
	partnerHoldsRoom = (noteGet(partnerReg, noteId(STOCK_OFFSET)) = "True")
end function

' nights between the two posted dates. an <input type="date"> posts
' yyyy-mm-dd, which parseDbDate reads
function nightsBetween(fromText, toText)
	dim d1, d2, diff
	nightsBetween = 0
	d1 = parseDbDate(fromText)
	d2 = parseDbDate(toText)
	if isempty(d1) then exit function
	if isempty(d2) then exit function
	diff = datediff("d", d1, d2)
	if diff > 0 then nightsBetween = diff
end function

' confprice holds one price per participant type for this conference
function confPrice(participantValue, otherValue)
	dim sqlp
	confPrice = 0
	if participantValue = "" then exit function
	if participantValue = "0" then exit function
	if not isnumeric(participantValue) then exit function
	' what an "other" participant type should cost has not been decided, so it
	' is left out of the total rather than guessed at
	if otherValue <> "0" and participantValue = otherValue then exit function
	sqlp = "select theprice, earlyprice, confdayprice "
	sqlp = sqlp & "from [imknasim].[dbo].[confprice] "
	sqlp = sqlp & "where theconference=" & theConf & " and theparticipant=" & clng(participantValue)
	r4.open sqlp,strconn,1,3
	if not r4.eof then
		confPrice = pickPrice(confTier, r4("earlyprice"), r4("theprice"), r4("confdayprice"))
	end if
	r4.close
end function

' every confhotel figure is per night: the room, the room when it is shared,
' and the breakfast supplement
function lodgingPrice(nights)
	dim roomId, nightly, sqlh
	lodgingPrice = 0
	if nights <= 0 then exit function
	if postedFlag(FLD_NEEDLODGING) <> "True" then exit function
	roomId = posted(FLD_ROOM)
	if roomId = "" then exit function
	if roomId = "0" then exit function
	if not isnumeric(roomId) then exit function

	' sharing a room does not change the rate - each occupant pays the rate
	' of the row they chose. withpartner only says the room may be shared
	sqlh = "select theprice, breakfast from [imknasim].[dbo].[confhotel] "
	sqlh = sqlh & "where theconference=" & theConf & " and theindex=" & clng(roomId)
	r4.open sqlh,strconn,1,3
	if not r4.eof then
		nightly = clng(n(r4("theprice")))
		if postedFlag(FLD_BREAKFAST) = "True" then nightly = nightly + clng(n(r4("breakfast")))
		lodgingPrice = nights * nightly
	end if
	r4.close
end function

theConf = trim(request.form("item") & "")
if theConf & "a" = "a" then theConf = "0"
if not isnumeric(theConf) then theConf = "0"

theReg = trim(request.form("reg") & "")
if not isnumeric(theReg) then theReg = ""
theCode = trim(request.form("code") & "")

sql = "select * from [imknasim].[dbo].[conference] where theIndex=" & theConf
r.open sql,strconn,1,3
confFound = not r.eof
confStart = ""
confClose = "0"
confEdit  = "False"
confShut  = "False"
confTier  = "regular"
editOpen  = false
if confFound then
	confStart = trim(r("startdate") & "")
	confClose = n(r("whentoclose"))
	confEdit  = b(r("isedit"))
	editOpen  = editIsOpen(r("isedit"), confStart, confClose)
	' closeform shuts the form by hand, whenclose shuts it on a date
	if formIsShut(r("closeform"), r("whenclose")) then confShut = "True"
	' the same tier getForm showed the visitor, decided here again so that
	' what is stored cannot differ from what was on screen
	confTier = priceTier(confStart, r("whenearly"))
end if
r.close

if confFound = false then
	response.write "{""ok"":""False"",""error"":""conference not found""}"
else
	' load the fields this conference actually has. anything posted that is not
	' on this list is ignored, so a crafted POST cannot write rows for other fields
	cnt = 0
	' decided once, because the mandatory check, the stored answer, the
	' stock guard and the pairing all have to agree about it
	shareOk = roomSharable()

	sql1 = "select f.theindex, f.fieldtype, f.ismandatory, f.choicetable, "
	sql1 = sql1 & "isnull(cf.ismandatory, 0) as confmand "
	sql1 = sql1 & "from conffield cf, fields f "
	sql1 = sql1 & "where f.theindex=cf.thefield and cf.theconf=" & theConf
	r1.open sql1,strconn,1,3
	while not r1.eof
		fId(cnt) = n(r1("theindex"))
		fType(cnt) = n(r1("fieldtype"))
		' mandatory at either level makes the field mandatory
		fMand(cnt) = "False"
		if b(r1("ismandatory")) = "True" then fMand(cnt) = "True"
		if b(r1("confmand")) = "True" then fMand(cnt) = "True"
		fChoice(cnt) = n(r1("choicetable"))
		fOther(cnt) = "0"
		if n(r1("fieldtype")) = "2" or n(r1("fieldtype")) = "7" then
			fOther(cnt) = otherOptionId(r1("choicetable"))
		end if
		' the roommate name rides on the same companion mechanism as "other",
		' but its owner is a yes/no field, so the value that opens it is True
		if fId(cnt) = cstr(FLD_PARTNER) and shareOk then fOther(cnt) = "True"
		if fId(cnt) = cstr(FLD_PARTICIPANT) then partOther = fOther(cnt)
		cnt = cnt + 1
		r1.movenext
	wend
	r1.close

	if partOther = "" then partOther = "0"

	' mandatory check. client side validation can be bypassed
	missing = ""
	for i = 0 to cnt - 1
		if fMand(i) = "True" then
			theVal = posted(fId(i))
			if fType(i) = "2" or fType(i) = "7" then
				if theVal = "0" then theVal = ""
			end if
			if isBlank(theVal) then missing = missing & fId(i) & ","
		end if
	next

	' a free text companion is required whenever its field is set to the
	' "other" option, otherwise the answer carries no information
	for i = 0 to cnt - 1
		if fOther(i) <> "0" then
			' a yes/no owner answers True, a choice owner answers an option id
			if answerOf(i) = fOther(i) then
				if isBlank(posted(OTHER_OFFSET + clng(fId(i)))) then
					missing = missing & (OTHER_OFFSET + clng(fId(i))) & ","
				end if
			end if
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
		isUpdate = false
		editClosed = false
		if theReg <> "" then
			isUpdate = true
			if editOpen = false then editClosed = true
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

		if confShut = "True" then
			response.write "{""ok"":""False"",""error"":""formclosed""}"
		elseif editClosed = true then
			response.write "{""ok"":""False"",""error"":""closed""}"
		elseif authOk = false then
			response.write "{""ok"":""False"",""error"":""badcode""}"
		else
			' the mobile must not already belong to another registration of this conference
			sql = "select theindex from [imknasim].[dbo].[registration] "
			sql = sql & "where theconf=" & theConf & " and cellnum='" & replace(theCell, "'", "''") & "'"
			if isUpdate = true then sql = sql & " and theindex<>" & theReg
			r.open sql,strconn,1,3
			taken = not r.eof
			r.close

			' a room may have sold out between loading the form and sending it.
			' the check is per field, so any field reading confhotel is covered
			noStock = false
			hotelField = "0"
			hotelNew = ""
			hotelOld = ""
			for i = 0 to cnt - 1
				if choiceTableName(fChoice(i)) = "confhotel" then
					hotelField = fId(i)
					hotelNew = posted(fId(i))
					if hotelNew = "0" then hotelNew = ""
					' a room is stock. a registrant who says no lodging is needed
					' holds none, whatever the selector was left on
					if postedFlag(FLD_NEEDLODGING) <> "True" then hotelNew = ""
				end if
			next
			if isUpdate = true and hotelField <> "0" then
				sql = "select thevalue from [imknasim].[dbo].[regfield] where thereg=" & theReg
				sql = sql & " and thefield=" & hotelField
				r.open sql,strconn,1,3
				if not r.eof then hotelOld = trim(r("thevalue") & "")
				r.close
			end if

			' what this registration was holding, and who it was sharing with,
			' before this save touched anything
			tookOld = false
			oldPartner = 0
			if isUpdate = true then
				if noteGet(theReg, noteId(STOCK_OFFSET)) = "True" then tookOld = true
				if isnumeric(noteGet(theReg, noteId(PAIR_OFFSET))) then
					oldPartner = clng(noteGet(theReg, noteId(PAIR_OFFSET)))
				end if
			end if
			' who this registrant is sharing with, and whether that person is
			' already holding the room. resolved before the stock guard, because
			' the second occupant of a room needs no free unit to join it
			partnerReg = 0
			partnerHolds = false
			if shareOk and postedFlag(FLD_PARTNER) = "True" then
				partnerReg = findPartner(theReg, hotelNew, posted(OTHER_OFFSET + FLD_PARTNER), trim(posted(FIRST_FIELD) & " " & posted(LAST_FIELD)))
				partnerHolds = partnerHoldsRoom(partnerReg, hotelNew)
			end if

			if hotelNew <> "" and hotelNew <> hotelOld and partnerHolds = false then
				if isnumeric(hotelNew) then
					sql = "select leftquantity from [imknasim].[dbo].[confhotel] where theindex=" & hotelNew
					r.open sql,strconn,1,3
					if r.eof then
						noStock = true
					elseif clng(n(r("leftquantity"))) <= 0 then
						noStock = true
					end if
					r.close
				end if
			end if

			if taken = true then
				response.write "{""ok"":""False"",""error"":""duplicate""}"
			elseif noStock = true then
				response.write "{""ok"":""False"",""error"":""nostock""}"
			else
				theFull = trim(posted(FIRST_FIELD) & " " & posted(LAST_FIELD))
				theMail = posted(EMAIL_FIELD)

				if isUpdate = true then
					sql = "select * from [imknasim].[dbo].[registration] where theindex=" & theReg
					r.open sql,strconn,1,3
					r("thename") = left(theFull, 255)
					r("cellnum") = left(theCell, 50)
					r("theemail") = left(theMail, 50)
					r.update
					outCode = trim(r("thecode") & "")
					r.close
				else
					if confEdit = "True" then
						outCode = newCode()
					else
						outCode = ""
					end if
					r.open "registration",strconn,1,3
					r.addnew
					r("theconf") = clng(theConf)
					r("thename") = left(theFull, 255)
					r("cellnum") = left(theCell, 50)
					r("theemail") = left(theMail, 50)
					r("thedate") = left(stampNow(), 10)
					r("thetime") = right(stampNow(), 5)
					if outCode <> "" then r("thecode") = outCode
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

				' the night count and the price are derived here rather than taken
				' from the browser, which only ever showed them read only
				srvNights = ""
				nightCount = nightsBetween(posted(FLD_CHECKIN), posted(FLD_CHECKOUT))
				if nightCount > 0 then srvNights = cstr(nightCount)

				srvPrice = ""
				totalPrice = confPrice(posted(FLD_PARTICIPANT), partOther) + lodgingPrice(nightCount)
				if totalPrice > 0 then srvPrice = cstr(totalPrice)

				for i = 0 to cnt - 1
					theVal = posted(fId(i))
					if fId(i) = cstr(FLD_NIGHTS) then theVal = srvNights
					if fId(i) = cstr(FLD_PRICE) then theVal = srvPrice
					if hotelField <> "0" and fId(i) = hotelField then theVal = hotelNew
					if fType(i) = "4" then theVal = postedFlag(fId(i))
					if fId(i) = cstr(FLD_PARTNER) and shareOk = false then theVal = "False"
					cntVals = 0
					if theVal <> "" then
						if fType(i) = "3" or fType(i) = "8" then
							' multi choice arrives comma separated, one row per selection
							parts = split(theVal, ",")
							for k = 0 to ubound(parts)
								onePart = trim(parts(k))
								if onePart <> "" then
									vals(cntVals) = onePart
									if fChoice(i) <> "0" then
										labels(cntVals) = choiceText(fChoice(i), onePart)
									else
										labels(cntVals) = confChoiceText(fId(i), onePart)
									end if
									cntVals = cntVals + 1
								end if
							next
						elseif fType(i) = "2" or fType(i) = "7" then
							if theVal <> "0" then
								vals(0) = theVal
								if fChoice(i) <> "0" then
									labels(0) = choiceText(fChoice(i), theVal)
								else
									labels(0) = confChoiceText(fId(i), theVal)
								end if
								cntVals = 1
							end if
						else
							vals(0) = theVal
							labels(0) = ""
							cntVals = 1
						end if
					end if
					call writeField(theReg, fId(i), cntVals)

					' the free text companion. writeField with a count of 0 removes any
					' row left over from a previous answer, so switching away from
					' "other" does not leave the old text behind
					if fOther(i) <> "0" then
						otherVal = ""
						if answerOf(i) = fOther(i) then otherVal = posted(OTHER_OFFSET + clng(fId(i)))
						if not isBlank(otherVal) then
							vals(0) = otherVal
							labels(0) = ""
							call writeField(theReg, OTHER_OFFSET + clng(fId(i)), 1)
						else
							call writeField(theReg, OTHER_OFFSET + clng(fId(i)), 0)
						end if
					end if
				next

				' hotel stock follows the saved answer. two registrants sharing a
				' room take one unit between them: the first takes it, the second
				' takes none. a note in regfield records which of the two this one is,
				' so an
				' edit returns only what was really taken. without a transaction a
				' rare double booking of the last room stays possible - the guard
				' before the save narrows the window
				tookNew = false
				if hotelNew <> "" and partnerHolds = false then tookNew = true

				releaseRoom = ""
				takeRoom = ""
				if tookOld = true then
					if tookNew = false or hotelNew <> hotelOld then releaseRoom = hotelOld
				end if
				if tookNew = true then
					if tookOld = false or hotelNew <> hotelOld then takeRoom = hotelNew
				end if

				' a room given up while the roommate still wants it does not go back
				' to stock - the holding passes to them instead
				handedOver = false
				if releaseRoom <> "" and oldPartner <> 0 then
					if roomStillWanted(oldPartner, releaseRoom) then
						call setHolding(oldPartner, true)
						handedOver = true
					end if
				end if

				if releaseRoom <> "" and handedOver = false and isnumeric(releaseRoom) then
					sql = "select * from [imknasim].[dbo].[confhotel] where theindex=" & releaseRoom
					r.open sql,strconn,1,3
					if not r.eof then
						r("leftquantity") = clng(n(r("leftquantity"))) + 1
						r.update
					end if
					r.close
				end if

				if takeRoom <> "" and isnumeric(takeRoom) then
					sql = "select * from [imknasim].[dbo].[confhotel] where theindex=" & takeRoom
					r.open sql,strconn,1,3
					if not r.eof then
						if clng(n(r("leftquantity"))) > 0 then
							r("leftquantity") = clng(n(r("leftquantity"))) - 1
							r.update
						end if
					end if
					r.close
				end if

				' record what this registration holds and who it shares with, and
				' keep the other side of the pairing in step
				call setHolding(theReg, tookNew)
				call setPartner(theReg, partnerReg)
				if oldPartner <> 0 and oldPartner <> partnerReg then
					call clearPartner(oldPartner, theReg)
				end if
				if partnerReg <> 0 then call setPartner(partnerReg, theReg)

				s = "{""ok"":""True"",""theNumber"":""" & theReg & ""","
				if isUpdate = true then
					s = s & """mode"":""edit""}"
				else
					s = s & """mode"":""new"",""theCode"":""" & outCode & ""","
					theDeadline = editDeadline(confStart, confClose)
					if confEdit = "True" and not isempty(theDeadline) then
						s = s & """editUntil"":""" & showDate(theDeadline) & """}"
					else
						s = s & """editUntil"":""""}"
					end if
				end if
				response.write s
			end if
		end if
	end if
end if
%>
