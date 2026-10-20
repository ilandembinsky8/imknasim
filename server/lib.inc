<%
' shared helpers for the registration endpoints.
' getForm.asp deliberately keeps its own copies - it is already live and
' there is no reason to risk it for the sake of de-duplication.

' country code assumed when the registrant types a local number with no prefix
const DEFAULT_COUNTRY = "972"

' failed code attempts allowed before the registration is locked, and for how long
const MAX_TRIES = 5
const LOCK_MINUTES = 15

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

' normalise a phone number to digits only, including the country code.
' a leading + or 00 means the country code is already there, otherwise a
' leading 0 is treated as local and replaced by the default country code.
'   050-1234567       -> 972501234567
'   +972 50 123 4567  -> 972501234567
'   00972501234567    -> 972501234567
'   +1 212 555 0101   -> 12125550101
function normCell(v)
	dim src, i, c, digits, hasCountry
	src = trim(cstr(v & ""))
	hasCountry = false
	if left(src, 1) = "+" then hasCountry = true

	digits = ""
	for i = 1 to len(src)
		c = mid(src, i, 1)
		if c >= "0" and c <= "9" then digits = digits & c
	next

	if digits = "" then
		normCell = ""
		exit function
	end if

	if left(digits, 2) = "00" then
		digits = mid(digits, 3)
		hasCountry = true
	end if

	if hasCountry = false then
		if left(digits, 1) = "0" then digits = DEFAULT_COUNTRY & mid(digits, 2)
	end if

	normCell = digits
end function

' timestamp as YYYY-MM-DD HH:MM. sorts correctly as a string, which is why
' the lock expiry fits in one nchar(16) column instead of a date plus a time
function stampAt(t)
	stampAt = year(t) & "-" & right("0" & month(t), 2) & "-" & right("0" & day(t), 2)
	stampAt = stampAt & " " & right("0" & hour(t), 2) & ":" & right("0" & minute(t), 2)
end function

function stampNow()
	stampNow = stampAt(now)
end function

' four digit code kept as text so a leading zero survives
function newCode()
	randomize
	newCode = right("000" & cstr(int(rnd * 10000)), 4)
end function

' parse a date string from the database into a real date.
' the admin has written three different shapes into the same column over time,
' so all three are accepted:
'   2026-10-14   padded iso, written by admin/add.asp
'   2027-6-14    unpadded iso, written by admin/edit.asp when the day is > 12
'   02/06/2027   dd/mm/yyyy, written by admin/edit.asp when the day is <= 12
' cdate() is deliberately avoided - it would read 02/06/2027 through the server
' locale and could return the 6th of February instead of the 2nd of June
function parseDbDate(v)
	dim t, parts, yy, mm, dd
	parseDbDate = empty
	t = trim(cstr(v & ""))
	if t = "" then exit function

	if instr(t, "/") > 0 then
		parts = split(t, "/")
		if ubound(parts) <> 2 then exit function
		dd = trim(parts(0))
		mm = trim(parts(1))
		yy = trim(parts(2))
	elseif instr(t, "-") > 0 then
		parts = split(t, "-")
		if ubound(parts) <> 2 then exit function
		yy = trim(parts(0))
		mm = trim(parts(1))
		dd = trim(parts(2))
	else
		exit function
	end if

	if not (isnumeric(yy) and isnumeric(mm) and isnumeric(dd)) then exit function
	yy = clng(yy)
	mm = clng(mm)
	dd = clng(dd)
	if yy < 1900 or yy > 2999 then exit function
	if mm < 1 or mm > 12 then exit function
	if dd < 1 or dd > 31 then exit function

	parseDbDate = DateSerial(yy, mm, dd)
end function

' a date back into the YYYY-MM-DD form the database stores
function dbDate(d)
	dbDate = year(d) & "-" & right("0" & month(d), 2) & "-" & right("0" & day(d), 2)
end function

' dd/mm/yyyy, for display
function showDate(d)
	showDate = right("0" & day(d), 2) & "/" & right("0" & month(d), 2) & "/" & year(d)
end function

' a conference day as the registrant sees it, for example: יום ד׳, 14/10/2026
function dayLabel(d)
	dim names
	names = Array("א׳", "ב׳", "ג׳", "ד׳", "ה׳", "ו׳", "שבת")
	dayLabel = "יום " & names(weekday(d) - 1) & ", " & showDate(d)
end function

' last date on which a registration may still be edited:
' the conference start date minus conference.whentoclose days
function editDeadline(startdateText, whenToClose)
	dim d
	d = parseDbDate(startdateText)
	if isempty(d) then
		editDeadline = empty
		exit function
	end if
	editDeadline = dateadd("d", -1 * clng(n(whenToClose)), d)
end function

' the single place that decides whether editing is allowed right now, so that
' getForm, getReg and saveForm can never disagree
function editIsOpen(isEditFlag, startdateText, whenToClose)
	dim d
	editIsOpen = false
	if b(isEditFlag) <> "True" then exit function
	d = editDeadline(startdateText, whenToClose)
	if isempty(d) then exit function
	if date() <= d then editIsOpen = true
end function

' what admin/edit.asp stores when a date box is saved empty - it substitutes
' "1/1/1" for a blank, which comes out of the formula as this date. nothing in
' the column distinguishes it from a real date, so it is read as "not set".
' this constant can go once edit.asp stops writing it
const EMPTY_DATE = "2001-01-01"

' trim() only strips spaces, so a tab or a newline survives it and makes an
' answer look filled in when there is nothing to read. every check for an
' empty answer goes through here
function isBlank(v)
	dim t
	t = cstr(v & "")
	t = replace(t, vbTab, " ")
	t = replace(t, vbCr, " ")
	t = replace(t, vbLf, " ")
	isBlank = (trim(t) = "")
end function

' a stored date that really carries a date, rather than a blank or the marker
function dateIsSet(v)
	dim t
	dateIsSet = false
	t = trim(cstr(v & ""))
	if t = "" then exit function
	if t = EMPTY_DATE then exit function
	if isempty(parseDbDate(t)) then exit function
	dateIsSet = true
end function

' the single place that decides whether the form still accepts answers.
' closeform shuts it by hand, whenclose shuts it on a date. the date itself is
' still open and the day after it is not, matching editIsOpen above
function formIsShut(closeFlag, whenCloseText)
	dim d
	formIsShut = false
	if b(closeFlag) = "True" then
		formIsShut = true
		exit function
	end if
	if not dateIsSet(whenCloseText) then exit function
	d = parseDbDate(whenCloseText)
	if date() > d then formIsShut = true
end function

' which of the three confprice columns applies today. the conference day is
' tested first, so a whenearly left later than the start date cannot keep the
' early price alive once the conference has already begun
function priceTier(startdateText, whenEarlyText)
	dim d
	priceTier = "regular"
	d = parseDbDate(startdateText)
	if not isempty(d) then
		if date() >= d then
			priceTier = "onsite"
			exit function
		end if
	end if
	if not dateIsSet(whenEarlyText) then exit function
	d = parseDbDate(whenEarlyText)
	if date() <= d then priceTier = "early"
end function

' the price of one tier. a tier the client never filled in falls back to the
' regular price - charging nothing would be a worse error than charging full
function pickPrice(tier, earlyPrice, regularPrice, onsitePrice)
	dim v
	v = 0
	if tier = "early" then v = clng(n(earlyPrice))
	if tier = "onsite" then v = clng(n(onsitePrice))
	if v <= 0 then v = clng(n(regularPrice))
	if v < 0 then v = 0
	pickPrice = v
end function
%>
