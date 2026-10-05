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
%>
