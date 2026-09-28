<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
theTable = request.form("theTable")
theNumber = request.form("theNumber")

if theNumber&"a"<>"a" then
	new1 = "no"
	sql = "select * from "&theTable&" where theindex="&theNumber
	r.open sql,strconn,1,3
else
	new1 = "yes"
	r.open theTable,strconn,1,3
	r.addnew
end if
if theTable="split" then
	r("theconference")=int(request.form("theConf"))
	r("theName")=request.form("theName")
	r("location")=request.form("location")
	r("theDay")=int(request.form("theDay"))
	r("theOrder")=int(request.form("theOrder"))
	r.update
end if
if theTable="session" then
	r("conference")=int(request.form("theConf"))
	r("theName")=request.form("theName")
	r("split")=int(request.form("split"))
	if request.form("showinagenda")="True" then
		r("showinagenda")=true
	else
		r("showinagenda")=false
	end if
	r("theOrder")=int(request.form("theOrder"))
	r("theDay")=int(request.form("theDay"))
	r.update
	r.close
	if new1="yes" then
		sql = "select * from "&theTable&" order by theIndex desc"
		r.open sql,strconn,1,3
		theNumber=r("theIndex")
		r.close
	end if
	if request.form("chairName")&"a"<>"a" then
		sql = "select * from chairman where theItem="&theNumber
		r.open sql,strconn,1,3
		if r.eof then
			r.close
			r.open "chairman",strconn,1,3
			r.addnew
			r("theItem")=theNumber
		end if
		r("theName")=request.form("chairName")
		r("theJob")=request.form("chairJob")
		r("theTitle")=int(request.form("chairTitle"))
		r("theOrganization")=int(request.form("chairOrganization"))
		r("theOrder")=1
		r.update
	end if
end if
if theTable="consponser" then
	r("conference")=int(request.form("theConf"))
	r("sponsor")=int(request.form("sponsor"))
	r("theOrder")=int(request.form("theOrder"))
	r.update
end if
if theTable="item" then
	r("conference")=int(request.form("theConf"))
	r("theName")=request.form("theName")
	r("session")=int(request.form("session"))
	r("sponsership")=int(request.form("sponsor"))
	r("itemType")=int(request.form("itemType"))
	r("theOrder")=int(request.form("theOrder"))
	r("fromHour")=request.form("fromHour")
	r("toHour")=request.form("toHour")
	r.update
	r.close
	if new1="yes" then
		sql = "select * from "&theTable&" order by theIndex desc"
		r.open sql,strconn,1,3
		theNumber=r("theIndex")
		r.close
	end if
	if request.form("lecName")&"a"<>"a" then
		sql = "select * from lecturer where theItem="&theNumber
		r.open sql,strconn,1,3
		if r.eof then
			r.close
			r.open "lecturer",strconn,1,3
			r.addnew
			r("theItem")=theNumber
		end if
		r("theName")=request.form("lecName")
		r("theJob")=request.form("lecJob")
		r("theTitle")=int(request.form("lecTitle"))
		r("theOrganization")=int(request.form("organization"))
		r("theOrder")=1
		r.update
	end if
end if
response.redirect "forumDetails.asp?num="&request.form("theConf")&"&item="&theTable&""
%>