
<!-- #include file="secure.inc" -->
<%
set r=Server.CreateObject("ADODB.Recordset")
set r1=Server.CreateObject("ADODB.Recordset")
set r2=Server.CreateObject("ADODB.Recordset")
if request("table")="split" then
	sql = "select * from [imknasim].[dbo].[split] where theIndex="&request("item")&""
	r.open sql,strconn,1,3
	if not r.eof then
		theName = replace(r("theName"),"'","''")
		s = "{""theNumber"":""" & r("theIndex") & """,""theName"":""" & theName & """,""location"":""" & r("location") & """,""theDay"":""" & r("theDay") & """,""theOrder"":""" & r("theOrder") & """}"
	else
		s = "{}"
	end if
	r.close
elseif request("table")="session" then
	sql = "select * from [imknasim].[dbo].[session] where theIndex="&request("item")&""
	r.open sql,strconn,1,3
	if not r.eof then
		theName = replace(r("theName"),"'","''")
		s = "{""theNumber"":""" & r("theIndex") & """,""theName"":""" & theName & """,""split"":""" & r("split") & """,""showinagenda"":""" & r("showinagenda") & """,""theOrder"":""" & r("theOrder") & """,""theDay"":""" & r("theDay") & """,""lecturers"":["
		sql1 = "select * from [imknasim].[dbo].[chairman] where theitem="&r("theIndex")
		firstItem=true
		r1.open sql1,strconn,1,3
		while not r1.eof
			if firstItem<>true then s = s + ","
			s = s & "{""theName"":"""&r1("theName")&""",""theJob"":"""&r1("theJob")&""",""theTitle"":"""&r1("theTitle")&""",""theOrganization"":"""&r1("theOrganization")&""",""theOrder"":"""&r1("theOrder")&"""}"
			firstItem=false
			r1.movenext
		wend
		r1.close
		s=s&"]}"
	else
		s = "{}"
	end if
	r.close
elseif request("table")="consponser" then
	sql = "select * from [imknasim].[dbo].[consponser] where theindex="&request("item")
	r.open sql,strconn,1,3
	if not r.eof then
		s = "{""theNumber"":""" & r("theIndex") & """,""sponsor"":""" & r("sponsor") & """,""theOrder"":""" & r("theOrder") & """}"
	else
		s = "{}"
	end if
	r.close
elseif request("table")="item" then
	sql = "select * from [imknasim].[dbo].[item] where theindex="&request("item")
	r.open sql,strconn,1,3
	if not r.eof then
		s = "{""theNumber"":""" & r("theIndex") & """,""theName"":""" & r("theName") & """,""session"":""" & r("session") & """,""sponsor"":""" & r("sponsership") & """,""theOrder"":""" & r("theOrder") & """,""itemType"":""" & r("itemType") & """,""fromHour"":""" & r("fromHour") & """,""toHour"":""" & r("toHour") & """,""lecturers"":["
		sql1 = "select * from [imknasim].[dbo].[lecturer] where theitem="&r("theIndex")
		firstItem=true
		r1.open sql1,strconn,1,3
		while not r1.eof
			if firstItem<>true then s = s + ","
			s = s & "{""theName"":"""&r1("theName")&""",""theJob"":"""&r1("theJob")&""",""theTitle"":"""&r1("theTitle")&""",""theOrganization"":"""&r1("theOrganization")&""",""theOrder"":"""&r1("theOrder")&"""}"
			firstItem=false
			r1.movenext
		wend
		r1.close
		s=s&"]}"
	else
		s = "{}"
	end if
	r.close
end if
response.write s
%>


